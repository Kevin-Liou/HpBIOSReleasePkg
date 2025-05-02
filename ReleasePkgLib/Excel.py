import os
import struct
import logging
import xlwings as xw
from colorama import Fore
from re import sub, search
from time import sleep
from xlwings import constants

from ReleasePkgLib import *
from .Platform import *

class ExcelChangeData:
    def __init__(self, platform_flag, match_folder_list):
        self.bios_fw_version = BiosBinaryFwVersion(platform_flag)
        self.BiosBuildDate = CheckBiosBuildDate(match_folder_list)

def CheckBiosBuildDate(Match_folder_list):
    BiosBuildDate = {}
    for Fv in Match_folder_list:
        time = ""
        if os.path.isdir(".\\" + Fv):
            Path = ".\\" + Fv + "\\Combined\\"
            for root,dirs,files in os.walk(Path):
                for name in files:
                    if ".inf" in name:
                        pattern = r"\d+/\d+/\d+"
                        File = open(root + "\\" + name, "r+")
                        Str_list = File.read()
                        searchObj = search(pattern, Str_list)
                        if searchObj:
                            time = searchObj.group(0)
                            File.close()
                            break
                        else:
                            File.close()
        if time == "":
            print(Fore.RED + "Check Build Date fail.")
        BiosBuildDate[Fv.split("_")[1]] = time
    return(BiosBuildDate)


def ModifyExcelData(Sheet, Modify_Data , Version):
    if not any(c.isdigit() or c.isnumeric() for c in Version):
        return
    Sheet_range = 'A1:A200'
    if Modify_Data == 'VERSION':
        logging.debug('VERSION')
        count = 0
        for char in Version:
            if char.isdigit():
                count += 1
        logging.debug('VERSION count=' + str(count))
        if count == 9:
            Sheet_range = 'B8:B12'
            logging.debug('ME Version:' + str(Version))
        elif (count == 6) or (count ==10):
            Sheet_range = 'B3:B7'
            logging.debug('BIOS Version:' + str(Version))

    if Modify_Data == 'ID':
        if Version == 'BIOS 000000':
            Sheet_range = 'B4'
            Version = '000000'
        elif Version == 'ME 000000':
            Sheet_range = 'B9'
            Version = '000000'

    if Modify_Data == 'PART NUMBER':
        if Version == 'BIOS P00000-000':
            Sheet_range = 'B3:B7'
            Version = 'P00000-000'
        elif Version == 'ME P00000-000':
            Sheet_range = 'B8:B12'
            Version = 'P00000-000'
    for row in Sheet.range(Sheet_range):
        for cell in row:
            if cell.value and Modify_Data in cell.value:
                if cell.value == Modify_Data:
                    logging.debug(cell.address + ' = '+ Modify_Data)
                    old_version = str(cell.offset(0, 1).value)
                    if Version in old_version:
                        logging.debug("The Version didn't need to update")
                        return
                    cell.offset(0, 1).value = "" + str(Version)
                    if old_version != cell.offset(0, 1).value :
                        cell.offset(0, 1).api.Font.Color = 0x00B050
                    if Modify_Data != 'System BIOS':
                        return


def FindOldMEVersion(Sheet):
    logging.debug('FindOldMEVersion')
    for row in Sheet.range('B8:B12'):
        for cell in row:
            if cell.value and 'VERSION' in cell.value:
                logging.debug('Old ME Version = ' + cell.offset(0, 1).value)
                return(cell.offset(0, 1).value)


def PrintBiosBuildDate(Match_folder_list, BiosBuildDate):
    for Fv in Match_folder_list:
        print(Fv.split("_")[1] + "_"+Fv.split("_")[2] + " Build Date:" + BiosBuildDate[Fv.split("_")[1]])


def CheckFileChecksum(Match_folder_list, NewVersion):
    try:
        BiosFileChecksum = {}
        for NProc in Match_folder_list:
            path = ".\\" + NProc
            if Platform_Flag(NProc) in Intel_Platforms:
                #======If Intel DM 400 16MB Binary
                if (os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin") and \
                    os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + ".bin")) or \
                    (os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_12.bin") and \
                    os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin")):
                    with open(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin", 'rb') as f:
                        content = f.read()
                        binary_sum = sum(bytearray(content))
                        binary_sum = hex(binary_sum & 0xFFFFFFFF)
                        f.close()
                    BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
                #======If Intel DM 800/600 32MB binary
                if os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin"):
                    with open(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin", 'rb') as f:
                        content = f.read()
                        binary_sum = sum(bytearray(content))
                        binary_sum = hex(binary_sum & 0xFFFFFFFF)
                        f.close()
                    BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
            #======If AMD
            else:
                if os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + ".bin") or \
                    os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin"):
                    for root,dirs,files in os.walk(path):
                        for name in files:
                            if "_16.bin" in name:
                                with open(path + "\\" + name, 'rb') as f:
                                    content = f.read()
                                    binary_sum = sum(bytearray(content))
                                    binary_sum = hex(binary_sum & 0xFFFFFFFF)
                                    f.close()
                                BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
                if os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin"):
                    with open(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin", 'rb') as f:
                        content = f.read()
                        binary_sum = sum(bytearray(content))
                        binary_sum = hex(binary_sum & 0xFFFFFFFF)
                        f.close()
                    BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
        logging.debug('BiosFileChecksum = ' + str(BiosFileChecksum))
        return(BiosFileChecksum)
    except:
        print(Fore.RED + "File Checksum Get Failed!")
        return 0


def PrintBiosBinaryChecksum(NewProcPkgInfo, BiosBinaryChecksum, NewVersion):
    for NProc in NewProcPkgInfo:
        path = ".\\" + ("_").join(NProc)
        logging.debug(str(BiosBinaryChecksum))
        #======If Intel
        if Platform_Flag(NProc) in Intel_Platforms:
            if os.path.isfile(path + "\\FPTW\\" + NProc[2] + "_" + NProc[3] + "_12.bin"):
                print(NProc[2] + "_" + NProc[3] + "_16.bin " + "checksum = 0x{}".format(BiosBinaryChecksum[NProc[2]].upper()))
            if os.path.isfile(path + "\\FPTW\\" + NProc[2] + "_" + NProc[3] + "_32.bin"):
                print(NProc[2] + "_" + NProc[3] + "_32.bin " + "checksum = 0x{}".format(BiosBinaryChecksum[NProc[2]].upper()))
        #======If AMD
        else:
            if (Platform_Flag(NProc) == "R26") or (Platform_Flag(NProc) == "Q26") or (Platform_Flag(NProc) == "Q27") or (Platform_Flag(NProc) == "S25"):
                print(NProc[0] + "_" + NewVersion + ".bin " + " checksum = 0x{}".format(BiosBinaryChecksum[NProc[0]].upper()))
            if Platform_Flag(NProc) == "R24" :
                print(NProc[1] + "_" + NewVersion + ".bin " + " checksum = 0x{}".format(BiosBinaryChecksum[NProc[1]].upper()))
            if (Platform_Flag(NProc) == "S27") or (Platform_Flag(NProc) == "S29") or (Platform_Flag(NProc) == "T25") or \
                (Platform_Flag(NProc) == "T26") or (Platform_Flag(NProc) == "T27"):
                print(NProc[0] + "_" + NewVersion + "_32.bin " + " checksum = 0x{}".format(BiosBinaryChecksum[NProc[0]].upper()))


def CheckMEVersion(NProc, Match_folder_list):
    Version = "11.0.11.1111" # Default version
    logging.debug('CheckMEVersion Start.')
    for Fv in Match_folder_list:
        if not os.path.isdir(".\\" + Fv):
            path = "..\\"
        else:
            path = ".\\"
        if Fv.split("_")[1] == NProc[2]:
            if os.path.isdir(path + Fv + "\\ME") and os.path.isfile(path + Fv + "\\ME\\ME.bin"):
                MEbinary_pattern = r"ME_+[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]+.bin"
                for root,dirs,files in os.walk(path + Fv + "\\ME"):
                    for name in files:
                        searchObj = search(MEbinary_pattern, name)
                        if searchObj != None:
                            ME_Version = name.split(".bin")[0] #Remove .bin
                            Version_List = ME_Version.split("_")
                            for Version in Version_List:
                                MEbinary_pattern = r"[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]"
                                searchObj = search(MEbinary_pattern, Version)
                                if searchObj != None:
                                    logging.debug('ME_Version1:' + Version)
                                    return Version
            elif Version == "11.0.11.1111" and os.path.isfile(path + Fv + "\\ME\\ME.inf"):
                pattern = r"DriverVer.+"
                File = open(path + Fv + "\\ME\\ME.inf","r+")
                Str_list = File.read()
                searchObj = search(pattern, Str_list)
                if searchObj != None:
                    Version = searchObj.group(0)[25:27] + ".0." + searchObj.group(0)[28:30] + "." + searchObj.group(0)[31:33] + searchObj.group(0)[34:36]
                    File.close()
                logging.debug('ME_Version2:' + Version)
                return Version
            else:
                return Version

def ModifyReleaseNote(NProc, ReleaseFileName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list):
    print("Platform ReleaseNote Modify...")
    app = xw.App(visible = False,add_book = False)
    app.display_alerts = False
    app.screen_updating = False
    filepath = r"{}".format(ReleaseFileName)
    check = "fail"
    logging.debug('app books open...')
    logging.debug("Platform_Flag = " + str(Platform_Flag(NProc)))
    print("Please wait for a moment while it is being processed...")
    wb = app.books.open(filepath) # The slow opening is a file problem.
    logging.debug('app books open finish.')
    #======For 1.0 Release note
    if "v1.08" in wb.sheets['Revision'].range('A8').value or \
    "v1.07" in wb.sheets['Revision'].range('A8').value or \
    "v1.06" in wb.sheets['Revision'].range('A8').value:
        try:
            #======If Intel DM G5 and late
            if Platform_Flag(ReleaseFileName) in Intel_Platforms_G5later:
                logging.debug('If Intel DM G5 and late')
                MEVersion = CheckMEVersion(NProc, Match_folder_list) # ex. 14.0.21.7227

                logging.debug('wb.sheets IntelProjectPN')
                IntelProjectPN = wb.sheets['IntelProjectPN']

                logging.debug('wb.sheets IntelPlatformInfo')
                IntelInfo = wb.sheets['IntelPlatformInfo']

                IntelHistory = None
                IntelHowToFlash = None
                # Check if 'IntelPlatformHistory_FY24' and 'IntelPlatformHowToFlash_FY24' exists
                for sheet in wb.sheets:
                    if sheet.name == 'IntelPlatformHistory_FY24':
                        IntelHistory = sheet
                        logging.debug(f'Found sheet: {sheet.name}')
                    if sheet.name == 'IntelPlatformHowToFlash_FY24':
                        IntelHowToFlash = sheet
                        logging.debug(f'Found sheet: {sheet.name}')

                # If can't find 'IntelPlatformHistory_FY24'，Check 'IntelPlatformHistory' exists
                if not IntelHistory:
                    for sheet in wb.sheets:
                        if 'IntelPlatformHistory' in sheet.name:
                            IntelHistory = sheet
                            logging.debug(f'Found sheet: {sheet.name}')
                            break

                # If can't find 'IntelPlatformHowToFlash_FY24'，Check 'IntelPlatformHowToFlash' exists
                if not IntelHowToFlash:
                    for sheet in wb.sheets:
                        if 'IntelPlatformHowToFlash' in sheet.name:
                            IntelHowToFlash = sheet
                            logging.debug(f'Found sheet: {sheet.name}')
                            break
                #======Marco work
                logging.debug('Marco work')
                IntelHistory.range('C:C').api.Insert(constants.InsertShiftDirection.xlShiftToRight) # Can't Protection worksheet
                CopyValues = IntelHistory.range('B9:B100').options(ndim=2).value
                IntelHistory.range('C9:C100').value = CopyValues
                IntelHistory.range('C9:C100').api.Font.Color = 0x000000
                logging.debug('Init finish.')
                #======Modify 'System BIOS Version' 'Build Date' 'CHECKSUM'
                logging.debug('Modify ,System BIOS ,Version ,Build Date ,CHECKSUM')
                check = ""
                if (NewBuildID == "" or NewBuildID == "0000"):
                    ModifyExcelData(IntelHistory, 'System BIOS Version', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]) # BIOS Version
                    ModifyExcelData(IntelProjectPN, 'VERSION', str(NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6])) # BIOS Version
                    ModifyExcelData(IntelHistory, 'BIOS Build Version', "0000") # BIOS Version
                else:
                    ModifyExcelData(IntelHistory, 'System BIOS Version', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID) # BIOS Version
                    ModifyExcelData(IntelProjectPN, 'System BIOS Version', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID) # BIOS Version
                    ModifyExcelData(IntelHistory, 'BIOS Build Version', NewBuildID) # NewBuildID
                ModifyExcelData(IntelHistory, 'Build Date', BiosBuildDate[ReleaseFileName.split("_")[2]]) # BUILD DATE
                ModifyExcelData(IntelHistory, 'CHECKSUM', "0x" + BiosBinaryChecksum[NProc[2]].upper()) # CHECK SUM
                ModifyExcelData(IntelHistory, 'MRC', BiosMrcVersion) # MRC
                if Platform_Flag(ReleaseFileName) in Intel_Platforms_G9later:
                    ModifyExcelData(IntelHistory, 'ISH FW version', "HpSigned_ishC_" + BiosIshVersion) # ISH
                else:
                    ModifyExcelData(IntelHistory, 'ISH FW version', "" + BiosIshVersion) # ISH
                ModifyExcelData(IntelHistory, 'PMC', BiosPmcVersion) # PMC
                ModifyExcelData(IntelHistory, 'NPHY', BiosNphyVersion) # NPHY
                ModifyExcelData(IntelProjectPN, 'PART NUMBER', "BIOS P00000-000") # BIOS PART NUMBER
                ModifyExcelData(IntelProjectPN, 'ID', "BIOS 000000") # BIOS PART NUMBER ID
                check = "pass"
                if check != "pass":
                    print("Can't find ['System BIOS Version', 'Target EE phase (DB/SI/PV)', 'Build Date', 'CHECKSUM']")
                #======ME Version
                logging.debug('ME Version')
                if not MEVersion == "11.0.11.1111":
                    pattern = r'[0-9]+[\.][0-9]+[\.][0-9]+[\.]\d{4}'
                    check = ""
                    OldMEVersion = FindOldMEVersion(IntelProjectPN)
                    logging.debug('ME Version : ' + MEVersion)
                    logging.debug('Old ME Version : ' + OldMEVersion)
                    ModifyExcelData(IntelProjectPN, 'VERSION', MEVersion)
                    if MEVersion != OldMEVersion:
                        ModifyExcelData(IntelProjectPN, 'PART NUMBER', "ME P00000-000") # ME PART NUMBER
                        ModifyExcelData(IntelProjectPN, 'ID', "ME 000000") # ME PART NUMBER ID
                    check = "pass"
                if check != "pass":
                    print("Can't find ['ME Firmware']")
                #======Modify 'Folder Path'
                check = ""
                for a in range(20, 40):
                    if IntelInfo.range('A'+str(a)).value == 'Folder Path' and \
                        IntelInfo.range('A'+str(a+1)).value == 'ODM FTP' and \
                        IntelInfo.range('A'+str(a+2)).value == 'Folder Path':
                        Folder_Path_HP = IntelInfo.range('B'+str(a)).value
                        Folder_Path_Quanta = IntelInfo.range('B'+str(a+2)).value
                        pattern = r'\w+_\w+_\w\d\d_\w+.7z'
                        Re_Folder_Path_HP = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_HP)
                        Re_Folder_Path_Quanta = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_Quanta)
                        IntelInfo.range('B'+str(a)).value = Re_Folder_Path_HP
                        IntelInfo.range('B'+str(a+2)).value = Re_Folder_Path_Quanta
                        logging.debug('Folder Path fill finish.')
                        check = "pass"
                        break
                if check != "pass":
                    print("Can't find ['Folder Path', 'ODM FTP', 'Folder Path']")
                #======Modify 'HowToFlash'
                check = ""
                for a in range(10, 35):
                    if IntelHowToFlash.range('A'+str(a)).value == "BIOS Flash: From -> To ":
                        IntelHowToFlash.range(str(a+2)+":"+str(a+2)).api.Insert()
                        CopyValues = IntelHowToFlash.range('A'+str(a+1)+':K'+str(a+1)).options(ndim = 2).value
                        IntelHowToFlash.range('A'+str(a+2)+':K'+str(a+2)).expand('table').value = CopyValues[0]
                        pattern = r'\d\d.\d\d.\d\d.*'
                        if ("->") not in str(IntelHowToFlash.range('A'+str(a+1)).value):
                            IntelHowToFlash.range('A'+str(a+1)).value = "00.00.00-> " + str(IntelHowToFlash.range('A'+str(a+1)).value)
                        Flash_Version_Left = sub(pattern, IntelHowToFlash.range('A'+str(a+1)).value.split("->")[1].strip(), IntelHowToFlash.range('A'+str(a+1)).value.split("->")[0])
                        if (NewBuildID == "" or NewBuildID == "0000"):
                            Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6], IntelHowToFlash.range('A'+str(a+1)).value.split("->")[1])
                        else:
                            Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]+"_"+NewBuildID, IntelHowToFlash.range('A'+str(a+1)).value.split("->")[1])
                        IntelHowToFlash.range('A'+str(a+1)).value = Flash_Version_Left + "->" + Flash_Version_Right
                        logging.debug('HowToFlash fill finish.')
                        check = "pass"
                        break
                if check != "pass":
                    print("Can't find ['BIOS Flash: From -> To ']")
                #======Modify 'HowToFlash' picture position
                flag = "0"
                for a in range(18, 100):
                    if IntelHowToFlash.range('E'+str(a)).value == "The User Account Control (UAC) setting needs to be set to DISABLED.":
                        pic_row = str(a-1)
                        flag = "1"
                        break
                logging.debug('Find image location.')
                if flag == "1" and len(IntelHowToFlash.pictures) == 1:
                    pic_name = IntelHowToFlash.pictures[0].name
                else:
                    flag == "0"
                logging.debug('Find image name. flag:' + flag)
                if flag == "1":
                    if IntelHowToFlash.pictures[0].name == pic_name:
                        i = 0
                        for i in range(10):
                            if (int(IntelHowToFlash.pictures[0].top) != int(IntelHowToFlash.range('E'+pic_row).top+13.00)):
                                #print(int(IntelHowToFlash.pictures[0].top), int(IntelHowToFlash.range('E'+pic_row).top+13.00))
                                IntelHowToFlash.pictures[0].top = IntelHowToFlash.range('E'+pic_row).top+13.00
                                sleep(0.5)
                                IntelHowToFlash.pictures[0].top = IntelHowToFlash.range('E'+pic_row).top+13.00
                                #print(int(IntelHowToFlash.pictures[0].top), int(IntelHowToFlash.range('E'+pic_row).top+13.00))
                else:
                    print("Unable to locate.\n")
                logging.debug('HowToFlash Pic position modify finish. \nModifyReleaseNote finish.')
            #======If AMD G5 and late DM
            elif Platform_Flag(ReleaseFileName) in AMD_Platforms_R24later:
                AMDHistory = wb.sheets['AMDPlatformHistory']
                AMDInfo = wb.sheets['AMDPlatformInfo']
                AMDHowToFlash = wb.sheets['AMDPlatformHowToFlash']
                #======Marco work
                AMDHistory.range('C:C').api.Insert(constants.InsertShiftDirection.xlShiftToRight)
                CopyValues = AMDHistory.range('B9:B100').options(ndim = 2).value
                AMDHistory.range('C9:C100').value = CopyValues
                AMDHistory.range('C9:C100').api.Font.Color = 0x000000
                logging.debug('Init finish.')
                #======Modify 'System BIOS Version' 'Build Date' 'CHECKSUM'
                check = ""
                for a in range(9, 30):
                    if AMDHistory.range('A'+str(a)).value == 'System BIOS Version' and \
                        AMDHistory.range('A'+str(a+1)).value == 'Target EE phase (DB/SI/PV)' and \
                        AMDHistory.range('A'+str(a+2)).value == 'Build Date' and \
                        AMDHistory.range('A'+str(a+3)).value == 'CHECKSUM':
                        if (NewBuildID == "" or NewBuildID == "0000"):
                            AMDHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
                        else:
                            AMDHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID
                        AMDHistory.range('B'+str(a+2)).value = BiosBuildDate[ReleaseFileName.split("_")[2]]
                        try :
                            AMDHistory.range('B'+str(a+3)).value = "0x" + BiosBinaryChecksum[NProc[2]].upper()
                        except :        # maybe some platform donot have .bin in root folder
                            AMDHistory.range('B'+str(a+3)).value = "0x"
                        logging.debug('Version fill finish.')
                        check = "pass"
                        break
                if check != "pass":
                    print("Can't find ['System BIOS Version', 'Target EE phase (DB/SI/PV)', 'Build Date', 'CHECKSUM']")
                #======Modify 'Folder Path'
                check = ""
                for a in range(20, 40):
                    if AMDInfo.range('A'+str(a)).value == 'Folder Path' and \
                        AMDInfo.range('A'+str(a+1)).value == 'ODM FTP' and \
                        AMDInfo.range('A'+str(a+2)).value == 'Folder Path':
                        Folder_Path_HP = AMDInfo.range('B'+str(a)).value
                        Folder_Path_Quanta = AMDInfo.range('B'+str(a+2)).value
                        if Platform_Flag(ReleaseFileName) == "R24":
                            pattern = r'\w+_\w\d\d_\d+.\d+.\d+.7z'
                        else:
                            pattern = r'\w\d\d_\w+.7z'
                        Re_Folder_Path_HP = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_HP)
                        Re_Folder_Path_Quanta = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_Quanta)
                        AMDInfo.range('B'+str(a)).value = Re_Folder_Path_HP
                        AMDInfo.range('B'+str(a+2)).value = Re_Folder_Path_Quanta
                        logging.debug('Folder Path fill finish.')
                        check = "pass"
                        break
                if check != "pass":
                    print("Can't find ['Folder Path', 'ODM FTP', 'Folder Path']")
                #======Modify 'HowToFlash'
                check = ""
                for a in range(10, 35):
                    if AMDHowToFlash.range('A'+str(a)).value == "BIOS Flash: From -> To ":
                        AMDHowToFlash.range(str(a+2)+':'+str(a+2)).api.Insert()
                        CopyValues = AMDHowToFlash.range('A'+str(a+1)+':K'+str(a+1)).options(ndim=2).value
                        AMDHowToFlash.range('A'+str(a+2)+':K'+str(a+2)).expand('table').value = CopyValues[0]
                        pattern = r'\d\d.\d\d.\d\d.*'
                        if Platform_Flag(ReleaseFileName) == "R24":
                            Flash_Version_Left = sub(pattern, AMDHowToFlash.range('A'+str(a+1)).value.split("->")[1].strip(), AMDHowToFlash.range('A'+str(a+1)).value.split(" ->")[0])
                        else:
                            Flash_Version_Left = sub(pattern, AMDHowToFlash.range('A'+str(a+1)).value.split("to")[1].strip(), AMDHowToFlash.range('A'+str(a+1)).value.split("to")[0])
                        if (NewBuildID == "" or NewBuildID == "0000"):
                            if Platform_Flag(ReleaseFileName) == "R24":
                                Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6], AMDHowToFlash.range('A'+str(a+1)).value.split(" ->")[1])
                            else:
                                Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6], AMDHowToFlash.range('A'+str(a+1)).value.split("to")[1])
                        else:
                            if Platform_Flag(ReleaseFileName) == "R24":
                                Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]+"_"+NewBuildID, AMDHowToFlash.range('A'+str(a+1)).value.split(" ->")[1])
                            else:
                                Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]+"_"+NewBuildID, AMDHowToFlash.range('A'+str(a+1)).value.split("to")[1])
                        if Platform_Flag(ReleaseFileName) == "R24":
                            AMDHowToFlash.range('A'+str(a+1)).value = Flash_Version_Left + " ->" + Flash_Version_Right
                        else:
                            AMDHowToFlash.range('A'+str(a+1)).value = Flash_Version_Left + "to" + Flash_Version_Right
                        logging.debug('HowToFlash fill finish. \nModifyReleaseNote finish.')
                        check = "pass"
                        break
                if check != "pass":
                    print("Can't find ['BIOS Flash: From -> To ']")
            wb.save()
            wb.close()
            app.quit()
            print("Platform ReleaseNote Modify " + Fore.GREEN + "succeeded.\n")
        except:
            wb.close()
            app.quit()
            print("Platform ReleaseNote Modify " + Fore.RED + "Failed!\n")
            return 0

    #======For 2.0 Release note
    elif ("v2.0" in wb.sheets['Revision'].range('A8').value) or ("v2.01.02" in wb.sheets['Revision'].range('A8').value):
        try:
            #======If Intel G9 and late
            if (Platform_Flag(ReleaseFileName) in Intel_Platforms_G9later) or (Platform_Flag(ReleaseFileName) in AMD_Platforms_G12later) :
                logging.debug('If Intel G9 later or AMD G12 later')
                MEVersion = CheckMEVersion(NProc, Match_folder_list) # ex. 14.0.21.7227
                logging.debug('wb.sheets ComponentInfo')
                ComponentInfo = wb.sheets['ComponentInfo']

                logging.debug('wb.sheets PlatformInfo')
                PlatformInfo = wb.sheets['PlatformInfo']

                logging.debug('wb.sheets PlatformHistory')
                PlatformHistory = wb.sheets['PlatformHistory']

                logging.debug('wb.sheets PlatformHowToFlash')
                PlatformHowToFlash = wb.sheets['PlatformHowToFlash']
                #======Marco work
                logging.debug('Marco work')
                PlatformHistory.range('C:C').api.Insert(constants.InsertShiftDirection.xlShiftToRight) # Can't Protection worksheet
                CopyValues = PlatformHistory.range('B9:B200').options(ndim=2).value
                PlatformHistory.range('C9:C200').value = CopyValues
                PlatformHistory.range('C9:C200').api.Font.Color = 0x000000
                logging.debug('Init finish.')
                #======Modify 'System BIOS Version' 'Build Date' 'CHECKSUM'
                logging.debug('Modify ,System BIOS ,Version ,Build Date ,CHECKSUM')
                check = ""
                if (NewBuildID == "" or NewBuildID == "0000"):
                    logging.debug("BuildID = "" or 0000")
                    ModifyExcelData(PlatformHistory, 'System BIOS Version', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]) # BIOS Version
                    ModifyExcelData(ComponentInfo, 'VERSION', str(NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6])) # BIOS Version
                    ModifyExcelData(PlatformHistory, 'BIOS Build Version', "0000") # BIOS Version
                else:
                    logging.debug("BuildID != 0000")
                    ModifyExcelData(PlatformHistory, 'System BIOS Version', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID) # BIOS Version
                    ModifyExcelData(ComponentInfo, 'VERSION', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID) # BIOS Version
                    ModifyExcelData(PlatformHistory, 'BIOS Build Version', NewBuildID) # NewBuildID
                BiosBuildDate = BiosBuildDate[NProc[2]]
                ModifyExcelData(PlatformHistory, 'Build Date', BiosBuildDate) # BUILD DATE
                ModifyExcelData(PlatformHistory, 'CHECKSUM', "0x" + BiosBinaryChecksum[NProc[2]].upper()) # CHECK SUM
                ModifyExcelData(PlatformHistory, 'MRC', BiosMrcVersion) # MRC
                ModifyExcelData(PlatformHistory, 'ISH FW version', "HpSigned_ishC_" + BiosIshVersion) # ISH
                ModifyExcelData(PlatformHistory, 'PMC', BiosPmcVersion) # PMC
                ModifyExcelData(PlatformHistory, 'NPHY FW  version', BiosNphyVersion) # NPHY
                ModifyExcelData(PlatformHistory, 'System BIOS', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + ".00") # System BIOS
                ModifyExcelData(PlatformHistory, 'HP System Firmware', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] +".00") # HP System Firmware
                ModifyExcelData(ComponentInfo, 'PART NUMBER', "BIOS P00000-000") # BIOS PART NUMBER
                ModifyExcelData(ComponentInfo, 'ID', "BIOS 000000") # BIOS PART NUMBER ID
                check = "pass"
                if check != "pass":
                    print("Can't find ['System BIOS Version', 'Target EE phase (DB/SI/PV)', 'Build Date', 'CHECKSUM']")
                #======ME Version
                logging.debug('ME Version')
                if not MEVersion == "11.0.11.1111":
                    pattern = r'[0-9]+[\.][0-9]+[\.][0-9]+[\.]\d{4}'
                    check = ""
                    OldMEVersion = FindOldMEVersion(ComponentInfo)
                    logging.debug('ME Version : ' + MEVersion)
                    logging.debug('Old ME Version : ' + OldMEVersion)
                    ModifyExcelData(ComponentInfo, 'VERSION', MEVersion)
                    if MEVersion != OldMEVersion:
                        ModifyExcelData(ComponentInfo,'PART NUMBER',"ME P00000-000") # ME PART NUMBER
                        ModifyExcelData(ComponentInfo,'ID',"ME 000000") # ME PART NUMBER ID
                    check = "pass"
                if check != "pass":
                    print("Can't find ['ME Firmware']")
                #======Modify 'Folder Path'
                logging.debug('Folder Path')
                check = ""
                for a in range(20, 40):
                    if PlatformInfo.range('A'+str(a)).value == 'Folder Path' and \
                        PlatformInfo.range('A'+str(a+1)).value == 'ODM FTP' and \
                        PlatformInfo.range('A'+str(a+2)).value == 'Folder Path':
                        Folder_Path_HP = PlatformInfo.range('B'+str(a)).value
                        Folder_Path_Quanta = PlatformInfo.range('B'+str(a+2)).value
                        pattern = r'\w+_\w+_\w\d\d_\w+.7z'
                        Re_Folder_Path_HP = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_HP)
                        Re_Folder_Path_Quanta = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_Quanta)
                        PlatformInfo.range('B'+str(a)).value = Re_Folder_Path_HP
                        PlatformInfo.range('B'+str(a+2)).value = Re_Folder_Path_Quanta
                        logging.debug('Folder Path fill finish.')
                        check = "pass"
                        break
                if check != "pass":
                    print("Can't find ['Folder Path', 'ODM FTP', 'Folder Path']")
                #======Modify 'HowToFlash'
                check = ""
                for a in range(10, 35):
                    if PlatformHowToFlash.range('A'+str(a)).value == "BIOS Flash: From -> To ":
                        PlatformHowToFlash.range(str(a+2)+":"+str(a+2)).api.Insert()
                        CopyValues = PlatformHowToFlash.range('A'+str(a+1)+':M'+str(a+1)).options(ndim = 2).value
                        PlatformHowToFlash.range('A'+str(a+2)+':K'+str(a+2)).expand('table').value = CopyValues[0]
                        pattern = r'\d\d.\d\d.\d\d.*'
                        if ("->") not in str(PlatformHowToFlash.range('A'+str(a+1)).value):
                            PlatformHowToFlash.range('A'+str(a+1)).value = "00.00.00-> " + str(PlatformHowToFlash.range('A'+str(a+1)).value)
                        Flash_Version_Left = sub(pattern, PlatformHowToFlash.range('A'+str(a+1)).value.split("->")[1].strip(), PlatformHowToFlash.range('A'+str(a+1)).value.split("->")[0])
                        if (NewBuildID == "" or NewBuildID == "0000"):
                            Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6], PlatformHowToFlash.range('A'+str(a+1)).value.split("->")[1])
                        else:
                            Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]+"_"+NewBuildID, PlatformHowToFlash.range('A'+str(a+1)).value.split("->")[1])
                        PlatformHowToFlash.range('A'+str(a+1)).value = Flash_Version_Left + "->" + Flash_Version_Right
                        logging.debug('HowToFlash fill finish.')
                        check = "pass"
                        break
                    elif PlatformHowToFlash.range('A'+str(a)).value == "From":
                        PlatformHowToFlash.range(str(a+2)+":"+str(a+2)).api.Insert()
                        CopyValues = PlatformHowToFlash.range('A'+str(a+1)+':O'+str(a+1)).options(ndim = 2).value
                        PlatformHowToFlash.range('A'+str(a+2)+':M'+str(a+2)).expand('table').value = CopyValues[0]
                        if ("From") not in str(PlatformHowToFlash.range('A'+str(a+1)).value):
                            PlatformHowToFlash.range('A'+str(a+1)).value = "00.00.00"
                            PlatformHowToFlash.range('B'+str(a+1)).value = "00.00.00"
                        PlatformHowToFlash.range('A'+str(a+1)).value = PlatformHowToFlash.range('B'+str(a+2)).value
                        if (NewBuildID == "" or NewBuildID == "0000"):
                            PlatformHowToFlash.range('B'+str(a+1)).value = NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]
                        else:
                            PlatformHowToFlash.range('B'+str(a+1)).value = NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]+"_"+NewBuildID
                        logging.debug('HowToFlash fill finish.')
                        check = "pass"
                        break
                if check != "pass":
                    print("Can't find ['BIOS Flash: From -> To ']")
                #======Modify 'HowToFlash' picture position
                flag = "0"
                for a in range(18, 100):
                    if PlatformHowToFlash.range('E'+str(a)).value == "The User Account Control (UAC) setting needs to be set to DISABLED.":
                        pic_row = str(a-1)
                        flag = "1"
                        break
                logging.debug('Find image location.')
                if flag == "1" and len(PlatformHowToFlash.pictures) == 1:
                    pic_name = PlatformHowToFlash.pictures[0].name
                else:
                    flag == "0"
                logging.debug('Find image name. flag:' + flag)
                if flag == "1":
                    if PlatformHowToFlash.pictures[0].name == pic_name:
                        i = 0
                        for i in range(10):
                            if (int(PlatformHowToFlash.pictures[0].top) != int(PlatformHowToFlash.range('E'+pic_row).top+13.00)):
                                #print(int(IntelHowToFlash.pictures[0].top), int(IntelHowToFlash.range('E'+pic_row).top+13.00))
                                PlatformHowToFlash.pictures[0].top = PlatformHowToFlash.range('E'+pic_row).top+13.00
                                sleep(0.5)
                                PlatformHowToFlash.pictures[0].top = PlatformHowToFlash.range('E'+pic_row).top+13.00
                                #print(int(IntelHowToFlash.pictures[0].top), int(IntelHowToFlash.range('E'+pic_row).top+13.00))
                else:
                    print("Unable to locate.\n")
                logging.debug('HowToFlash Pic position modify finish. \nModifyReleaseNote finish.')
            #======If AMD for G12 (Not use now)
            # elif Platform_Flag(ReleaseFileName) in AMD_Platforms_R24later:
            #     AMDHistory = wb.sheets['AMDPlatformHistory']
            #     AMDInfo = wb.sheets['AMDPlatformInfo']
            #     AMDHowToFlash = wb.sheets['AMDPlatformHowToFlash']
            #     #======Marco work
            #     AMDHistory.range('C:C').api.Insert(constants.InsertShiftDirection.xlShiftToRight)
            #     CopyValues = AMDHistory.range('B9:B100').options(ndim = 2).value
            #     AMDHistory.range('C9:C100').value = CopyValues
            #     AMDHistory.range('C9:C100').api.Font.Color = 0x000000
            #     logging.debug('Init finish.')
            #     #======Modify 'System BIOS Version' 'Build Date' 'CHECKSUM'
            #     check = ""
            #     for a in range(9, 30):
            #         if AMDHistory.range('A'+str(a)).value == 'System BIOS Version' and \
            #             AMDHistory.range('A'+str(a+1)).value == 'Target EE phase (DB/SI/PV)' and \
            #             AMDHistory.range('A'+str(a+2)).value == 'Build Date' and \
            #             AMDHistory.range('A'+str(a+3)).value == 'CHECKSUM':
            #             if (NewBuildID == "" or NewBuildID == "0000"):
            #                 AMDHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
            #             else:
            #                 AMDHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID
            #             AMDHistory.range('B'+str(a+2)).value = BiosBuildDate[ReleaseFileName.split("_")[2]]
            #             try :
            #                 AMDHistory.range('B'+str(a+3)).value = "0x" + BiosBinaryChecksum[NProc[2]].upper()
            #             except :        # maybe some platform donot have .bin in root folder
            #                 AMDHistory.range('B'+str(a+3)).value = "0x"
            #             logging.debug('Version fill finish.')
            #             check = "pass"
            #             break
            #     if not check == "pass":
            #         print("Can't find ['System BIOS Version', 'Target EE phase (DB/SI/PV)', 'Build Date', 'CHECKSUM']")
            #     #======Modify 'Folder Path'
            #     check = ""
            #     for a in range(20, 40):
            #         if AMDInfo.range('A'+str(a)).value == 'Folder Path' and \
            #             AMDInfo.range('A'+str(a+1)).value == 'ODM FTP' and \
            #             AMDInfo.range('A'+str(a+2)).value == 'Folder Path':
            #             Folder_Path_HP = AMDInfo.range('B'+str(a)).value
            #             Folder_Path_Quanta = AMDInfo.range('B'+str(a+2)).value
            #             if Platform_Flag(ReleaseFileName) == "R24":
            #                 pattern = r'\w+_\w\d\d_\d+.\d+.\d+.7z'
            #             else:
            #                 pattern = r'\w\d\d_\w+.7z'
            #             Re_Folder_Path_HP = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_HP)
            #             Re_Folder_Path_Quanta = sub(pattern, ("_").join(NProc) + ".7z", Folder_Path_Quanta)
            #             AMDInfo.range('B'+str(a)).value = Re_Folder_Path_HP
            #             AMDInfo.range('B'+str(a+2)).value = Re_Folder_Path_Quanta
            #             logging.debug('Folder Path fill finish.')
            #             check = "pass"
            #             break
            #     if not check == "pass":
            #         print("Can't find ['Folder Path', 'ODM FTP', 'Folder Path']")
            #     #======Modify 'HowToFlash'
            #     check = ""
            #     for a in range(10, 35):
            #         if AMDHowToFlash.range('A'+str(a)).value == "BIOS Flash: From -> To ":
            #             AMDHowToFlash.range(str(a+2)+':'+str(a+2)).api.Insert()
            #             CopyValues = AMDHowToFlash.range('A'+str(a+1)+':K'+str(a+1)).options(ndim=2).value
            #             AMDHowToFlash.range('A'+str(a+2)+':K'+str(a+2)).expand('table').value = CopyValues[0]
            #             pattern = r'\d\d.\d\d.\d\d.*'
            #             if Platform_Flag(ReleaseFileName) == "R24":
            #                 Flash_Version_Left = sub(pattern, AMDHowToFlash.range('A'+str(a+1)).value.split("->")[1].strip(), AMDHowToFlash.range('A'+str(a+1)).value.split(" ->")[0])
            #             else:
            #                 Flash_Version_Left = sub(pattern, AMDHowToFlash.range('A'+str(a+1)).value.split("to")[1].strip(), AMDHowToFlash.range('A'+str(a+1)).value.split("to")[0])
            #             if (NewBuildID == "" or NewBuildID == "0000"):
            #                 if Platform_Flag(ReleaseFileName) == "R24":
            #                     Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6], AMDHowToFlash.range('A'+str(a+1)).value.split(" ->")[1])
            #                 else:
            #                     Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6], AMDHowToFlash.range('A'+str(a+1)).value.split("to")[1])
            #             else:
            #                 if Platform_Flag(ReleaseFileName) == "R24":
            #                     Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]+"_"+NewBuildID, AMDHowToFlash.range('A'+str(a+1)).value.split(" ->")[1])
            #                 else:
            #                     Flash_Version_Right = sub(pattern, NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]+"_"+NewBuildID, AMDHowToFlash.range('A'+str(a+1)).value.split("to")[1])
            #             if Platform_Flag(ReleaseFileName) == "R24":
            #                 AMDHowToFlash.range('A'+str(a+1)).value = Flash_Version_Left + " ->" + Flash_Version_Right
            #             else:
            #                 AMDHowToFlash.range('A'+str(a+1)).value = Flash_Version_Left + "to" + Flash_Version_Right
            #             logging.debug('HowToFlash fill finish. \nModifyReleaseNote finish.')
            #             check = "pass"
            #             break
            #     if not check == "pass":
            #         print("Can't find ['BIOS Flash: From -> To ']")
            wb.save()
            wb.close()
            app.quit()
            print("Platform ReleaseNote Modify " + Fore.GREEN + "succeeded.\n")
        except:
            wb.close()
            app.quit()
            print("Platform ReleaseNote Modify " + Fore.RED + "Failed!\n")
            return 0

    else:
        Old_ReleaseNoteVersion = wb.sheets['Revision'].range('A8').value
        wb.close()
        app.quit()
        print("ReleaseNote Version is " + Old_ReleaseNoteVersion + " Not in 1.06~1.08 Version, Modify Skip.\n")