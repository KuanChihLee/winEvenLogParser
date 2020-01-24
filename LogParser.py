import os
import wx
import re
import json
import glob
import qrcode
import pprint
from collections import deque
from datetime import datetime
import Evtx.Evtx as evtx
import Evtx.Views as e_views
import xml.etree.ElementTree as ET

eventFolder = './pythonApp/logs/*.evtx'
QR_path = './pythonApp/Qrcode/'
errorBook_dir = './pythonApp/errorCode/'

class Panel(wx.Panel):
    def __init__(self, parent, path, image=True, text=''):
        super(Panel, self).__init__(parent, -1)
        if image:
            self.bitmap = wx.Bitmap(path)
            self.scale_bitmap(self.bitmap, 400, 400)
            control = wx.StaticBitmap(self, -1, self.bitmap)
            # control.SetPosition((20, 20))
        else:
            control = wx.StaticText(self, -1, text, pos=(20, 20))

    def scale_bitmap(self, bitmap, width, height):
        image = wx.Bitmap.ConvertToImage(bitmap)
        image = image.Scale(width, height, wx.IMAGE_QUALITY_HIGH)
        self.bitmap = wx.Bitmap(image)

class logParser:
    def __init__(self):
        self.counter = 0
        self.period_min = 5
        self.period_max = 60
        self.record_dict = {}
        self.feature_dict = {}
        self.version = '1.0.0'
        self.feature_dict['version'] = self.version
        self.root_tag_pattern = re.compile(r'(\{.*\})(.*)')
        self.str_block_pattern = re.compile(r'<string>\s*(\S.*)\s*</string>\n*')
        self.error_code_pattern = re.compile(r'0x20001(?P<base>[0-3]{1})(?P<hex>.{2})')
        self.error_context_pattern = re.compile(r'.*EW:\s*(0x.*)\s*=\s*(\S.*\S)\s*<.*>.*')
        self.error_clear_pattern = re.compile(r'.*Fault Cleared\:\s*(\S.*\S)\s*')
        self.error_set_pattern = re.compile(r'.*Fault Set\:\s*(\S.*\S)\s*')
        self.first_error_pattern = re.compile(r'.*(Device fault detected in  Mass Spectrometer).*')
        self.errorCode_table = {}
        try:
            with open(errorBook_dir+'Table.json', 'r') as fp:
                self.errorCode_table = json.load(fp)
        except EnvironmentError: # parent of IOError, OSError *and* WindowsError where available
            print("No troubleshooting table found!")

    def __rootChildren(self, treeRoot):
        root_xmlns = self.root_tag_pattern.search(treeRoot.tag)[1]
        eventdata = treeRoot.find(root_xmlns+'EventData')
        system = treeRoot.find(root_xmlns+'System')
        return root_xmlns, eventdata, system

    def __systemMembers(self, system, root_xmlns):
        provider = system.find(root_xmlns+'Provider')
        provider_name = list(provider.attrib.values())[0]

        timecreated = system.find(root_xmlns+'TimeCreated')
        timestamp = list(timecreated.attrib.values())[0]

        computer = system.find(root_xmlns+'Computer').text
        keywords = system.find(root_xmlns+'Keywords').text
        security = system.find(root_xmlns+'Security')
        userId = list(security.attrib.values())[0]
        return provider_name, timestamp, computer, keywords, userId

    def __error_code_pair_match(self, code, base, hexNum, errorText):
        # In general, because it's backward reading, we want to find Failure Cleared in previous Errorcode
        # So, current Errorcode is Failure Set
        cur_fault_set_match = self.error_set_pattern.search(errorText)
        if cur_fault_set_match is not None:
            base = int(base)
            record_decimal = int(hexNum, 16)
            delFlag = False
            for prevCode in self.record_dict.keys():
                sub_match = self.error_code_pattern.search(prevCode)
                prev_code = sub_match.group(0)
                prev_base = int(sub_match.group("base"))
                prev_decimal = int(sub_match.group("hex"), 16)

                if base == prev_base:
                    if base == 0:   # base = 0: API Based Error
                        # api dictionary exits or not
                        if self.errorCode_table:
                            # Make sure Errorcode in Table
                            if prevCode in self.errorCode_table:
                                search_result = self.errorCode_table[prevCode]
                                # Make sure this Errorcode is Fault Cleared text
                                fault_clear_match = self.error_clear_pattern.search(search_result)
                                if fault_clear_match is not None:
                                    # Check the main content in Fault Set and Fault Cleared is same text
                                    if cur_fault_set_match.group(1) == fault_clear_match.group(1):
                                        delFlag = True
                                        break
                        else:
                            break
                    elif base == 1: # base = 1: LCS Based Error
                        if prev_decimal-16 == record_decimal:
                            delFlag = True
                            break
                    elif base == 3: # base = 3: VPS Based Error
                        # vps dictionary exits or not
                        if self.errorCode_table:
                            # Make sure Errorcode in Table
                            if prevCode in self.errorCode_table:
                                search_result = self.errorCode_table[prevCode]
                                # Make sure this Errorcode is Fault Cleared text
                                fault_clear_match = self.error_clear_pattern.search(search_result)
                                if fault_clear_match is not None:
                                    # Check the main content in Fault Set and Fault Cleared is same text
                                    if cur_fault_set_match.group(1) == fault_clear_match.group(1):
                                        delFlag = True
                                        break
                        else:
                            break
            if delFlag:
                return prev_code
            else:
                return None

    def getLastSigError(self, record):
        root = ET.fromstring(record)
        root_xmlns, eventdata, system = self.__rootChildren(root)
        # Find error text match the pattern
        metadata = eventdata.find(root_xmlns+'Data').text
        first_error_match = self.first_error_pattern.search(metadata)
        # Check Provider is 'Analyst'
        provider, timestamp, computer, keywords, userId = self.__systemMembers(system, root_xmlns)
        if first_error_match is None or provider != 'Analyst':
            return '', False
        else:
            # Record the first error information/details
            self.feature_dict["Error"] = first_error_match.group(1)
            self.feature_dict["TimeCreated"] = timestamp
            self.feature_dict["Computer"] = computer
            self.feature_dict["Keywords"] = keywords
            self.feature_dict["UserId"] = userId
            return timestamp, True

    def getSigDetails(self, record, date):
        def __update_feature_dict():
            for key, value in self.record_dict.items():
                error_clear_match = self.error_clear_pattern.search(value["Description"])
                if error_clear_match is None:
                    self.feature_dict[key] = 1
                    self.feature_dict[(key+' Description')] = value["Description"]
                    del value["Description"]
                    self.feature_dict[(key+' Details')] = value
            if not self.errorCode_table:
                self.feature_dict['Warning'] = "Can't find Troubleshooting Spreadsheet"

        root = ET.fromstring(record)
        root_xmlns, eventdata, system = self.__rootChildren(root)
        metadata = eventdata.find(root_xmlns+'Data').text

        # If another error happens and it's equal/similar to the first error we found before, record it.
        first_error_match = self.first_error_pattern.search(metadata)
        if first_error_match is not None:
            if 'Error_repeat' not in self.feature_dict:
                self.feature_dict["Error_repeat"] = 1
            else:
                self.feature_dict["Error_repeat"] += 1
        
        provider, cur_timestamp, computer, keywords, userId = self.__systemMembers(system, root_xmlns)
        cur_datetime = datetime.strptime(cur_timestamp, '%Y-%m-%d %H:%M:%S')
        err_datetime = datetime.strptime(date, '%Y-%m-%d %H:%M:%S')
        if provider != 'Analyst':
            self.counter += 1
            if self.counter > 10:
                diff = abs(cur_datetime - err_datetime)
                if diff.days < 1:
                    if diff.seconds > 60:
                        __update_feature_dict()
                        return True
                    else:
                        self.counter = 0
                else:
                    __update_feature_dict()
                    return True
        else:
            self.counter = 0
            diff = abs(cur_datetime - err_datetime)
            if diff.days >= 1:
                __update_feature_dict()
                return True
            else:
                if diff.seconds > self.period_max:
                    __update_feature_dict()
                    return True

                if diff.seconds > self.period_min:
                    # It's error code pattern
                    err_context_match = self.error_context_pattern.search(metadata)
                    if err_context_match is not None:
                        # To see error code: 0x2...
                        error_code_match = self.error_code_pattern.search(err_context_match.group(1))
                        error_code_text = err_context_match.group(2)
                        if error_code_match is not None:
                            errorCode = error_code_match.group(0)
                            base = error_code_match.group(1)
                            hexNum = error_code_match.group(2)
                            if error_code_match.group(0) not in self.record_dict:
                                data_dict = {}
                                datastring = ''
                                data_dict["TimeBefore"] = diff.seconds
                                data_dict["Computer"] = computer
                                data_dict["Keywords"] = keywords
                                data_dict["UserId"] = userId
                                for data in re.finditer(self.str_block_pattern, metadata):
                                    datastring += data.group(1) + ';'
                                data_dict["Metadata"] = datastring
                                data_dict["Description"] = err_context_match.group(2)
                                # Times of repeated error
                                data_dict["Repeat"] = 0
                                # Default: can be searched in Troubleshooting Table (spreadsheet)
                                # data_dict["InTable"] = 1
                                self.record_dict[errorCode] = data_dict
                            else:
                                self.record_dict[errorCode]["Repeat"] += 1

                            prevCode = self.__error_code_pair_match(errorCode, base, hexNum, error_code_text)
                            if prevCode:
                                if self.record_dict[error_code_match.group(0)]["Repeat"] > 0:
                                    self.record_dict[error_code_match.group(0)]["Repeat"] -= 1
                                else:
                                    del self.record_dict[error_code_match.group(0)]
                                if self.record_dict[prevCode]["Repeat"] > 0:
                                    self.record_dict[prevCode]["Repeat"] -= 1
                                else:
                                    del self.record_dict[prevCode]
        return False

if __name__ == '__main__':
    app = wx.App()
    frame = wx.Frame(None, -1, 'QRcode')
    frame.SetSize(450,450)
    dw, dh = wx.DisplaySize()
    frame.SetPosition((dw/3, dh/4))

    logfiles = glob.glob(eventFolder) # * means all if need specific format then *.evtx
    if len(logfiles) > 0:
        latest_file = max(logfiles, key=os.path.getmtime)
        file_name = os.path.basename(latest_file)[:-5]

        findErr = False
        err_timestamp = ''
        logparser = logParser()
        with evtx.Evtx(latest_file) as log:
            dequeLog = deque(log.records())
            while dequeLog:
                last_record = dequeLog.pop()
                if findErr:
                    if logparser.getSigDetails(last_record.xml(), err_timestamp):
                        break
                else:
                    err_timestamp, findErr = logparser.getLastSigError(last_record.xml())
        
        # pp = pprint.PrettyPrinter()
        if logparser.feature_dict:
            # pp.pprint(logparser.feature_dict)
            feature_json = json.dumps(logparser.feature_dict)
            qr = qrcode.make(feature_json)
            file_dir = os.path.dirname(QR_path)
            if not os.path.exists(file_dir):
                os.makedirs(file_dir)
            qr.save(QR_path + file_name + '_QRcode.png')
            panel = Panel(frame, QR_path + file_name + '_QRcode.png')
        else:
            panel = Panel(frame, './', False, "No Error level found, No QRcode")
    else:
        panel = Panel(frame, './', False, "No log file exists, No QRcode")

    frame.Show()
    app.MainLoop()