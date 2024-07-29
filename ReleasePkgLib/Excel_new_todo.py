# ==============================================================================
# 新的模組化 Excel 修改方式, 還需要驗證
# ==============================================================================


import os
import logging
import xlwings as xw
from colorama import Fore
from re import sub, search
from time import sleep
from xlwings import constants

from ReleasePkgLib import *
from .Platform import *

# 檢查 BIOS 編譯日期
def CheckBiosBuildDate(Match_folder_list):
    BiosBuildDate = {}
    for Fv in Match_folder_list:
        time = ""
        if os.path.isdir(".\\" + Fv):
            Path = ".\\" + Fv + "\\Combined\\"
            for root, dirs, files in os.walk(Path):
                for name in files:
                    if ".inf" in name:
                        pattern = r"\d+/\d+/\d+"
                        with open(root + "\\" + name, "r+") as file:
                            Str_list = file.read()
                            searchObj = search(pattern, Str_list)
                            if searchObj:
                                time = searchObj.group(0)
                                break
        if time == "":
            print(Fore.RED + "Check Build Date fail.")
        BiosBuildDate[Fv.split("_")[1]] = time
    return BiosBuildDate

# 修改 Excel 資料
def Modify_Excel_Data(Sheet, Modify_Data, Version):
    if not any(c.isdigit() or c.isnumeric() for c in Version):
        return
    Sheet_range = 'A1:A200'
    if Modify_Data == 'VERSION':
        count = sum(char.isdigit() for char in Version)
        if count == 9:
            Sheet_range = 'B8:B12'
        elif count in (6, 10):
            Sheet_range = 'B3:B7'

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
                    old_version = str(cell.offset(0, 1).value)
                    if Version in old_version:
                        logging.debug("The Version didn't need to update")
                        return
                    cell.offset(0, 1).value = "" + str(Version)
                    if old_version != cell.offset(0, 1).value:
                        cell.offset(0, 1).api.Font.Color = 0x00B050
                    if Modify_Data != 'System BIOS':
                        return

# 查找舊的 ME 版本
def Find_Old_ME_Version(Sheet):
    logging.debug('FindOldMEVersion')
    for row in Sheet.range('B8:B12'):
        for cell in row:
            if cell.value and 'VERSION' in cell.value:
                return cell.offset(0, 1).value

# 打印 BIOS 編譯日期
def PrintBiosBuildDate(Match_folder_list, BiosBuildDate):
    for Fv in Match_folder_list:
        print(Fv.split("_")[1] + "_" + Fv.split("_")[2] + " Build Date:" + BiosBuildDate[Fv.split("_")[1]])

# 檢查文件校驗和
def CheckFileChecksum(Match_folder_list, NewVersion):
    try:
        BiosFileChecksum = {}
        for NProc in Match_folder_list:
            path = ".\\" + NProc
            if Platform_Flag(NProc) in Intel_Platforms:
                # Intel DM 400 16MB Binary
                if (os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin") and \
                    os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + ".bin")) or \
                    (os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_12.bin") and \
                    os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin")):
                    with open(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin", 'rb') as f:
                        content = f.read()
                        binary_sum = sum(bytearray(content))
                        binary_sum = hex(binary_sum & 0xFFFFFFFF)
                    BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
                # Intel DM 800/600 32MB binary
                if os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin"):
                    with open(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin", 'rb') as f:
                        content = f.read()
                        binary_sum = sum(bytearray(content))
                        binary_sum = hex(binary_sum & 0xFFFFFFFF)
                    BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
            # AMD
            else:
                if os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + ".bin") or \
                    os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_16.bin"):
                    for root, dirs, files in os.walk(path):
                        for name in files:
                            if "_16.bin" in name:
                                with open(path + "\\" + name, 'rb') as f:
                                    content = f.read()
                                    binary_sum = sum(bytearray(content))
                                    binary_sum = hex(binary_sum & 0xFFFFFFFF)
                                BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
                if os.path.isfile(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin"):
                    with open(path + "\\" + NProc.split("_")[1] + "_" + NewVersion + "_32.bin", 'rb') as f:
                        content = f.read()
                        binary_sum = sum(bytearray(content))
                        binary_sum = hex(binary_sum & 0xFFFFFFFF)
                    BiosFileChecksum[NProc.split("_")[1]] = binary_sum[2:]
        logging.debug('BiosFileChecksum = ' + str(BiosFileChecksum))
        return BiosFileChecksum
    except Exception as e:
        logging.error(f"File Checksum Get Failed: {e}")
        print(Fore.RED + "File Checksum Get Failed!")
        return 0

# 打印 BIOS 二進制校驗和
def PrintBiosBinaryChecksum(NewProcPkgInfo, BiosBinaryChecksum, NewVersion):
    for NProc in NewProcPkgInfo:
        path = ".\\" + ("_").join(NProc)
        logging.debug(str(BiosBinaryChecksum))
        if Platform_Flag(NProc) in Intel_Platforms:
            if os.path.isfile(path + "\\FPTW\\" + NProc[2] + "_" + NProc[3] + "_12.bin"):
                print(NProc[2] + "_" + NProc[3] + "_16.bin " + "checksum = 0x{}".format(BiosBinaryChecksum[NProc[2]].upper()))
            if os.path.isfile(path + "\\FPTW\\" + NProc[2] + "_" + NProc[3] + "_32.bin"):
                print(NProc[2] + "_" + NProc[3] + "_32.bin " + "checksum = 0x{}".format(BiosBinaryChecksum[NProc[2]].upper()))
        else:
            if Platform_Flag(NProc) in ["R26", "Q26", "Q27", "S25"]:
                print(NProc[0] + "_" + NewVersion + ".bin " + " checksum = 0x{}".format(BiosBinaryChecksum[NProc[0]].upper()))
            if Platform_Flag(NProc) == "R24":
                print(NProc[1] + "_" + NewVersion + ".bin " + " checksum = 0x{}".format(BiosBinaryChecksum[NProc[1]].upper()))
            if Platform_Flag(NProc) in ["S27", "S29", "T25", "T26", "T27"]:
                print(NProc[0] + "_" + NewVersion + "_32.bin " + " checksum = 0x{}".format(BiosBinaryChecksum[NProc[0]].upper()))

# 檢查 ME 版本
def CheckMEVersion(NProc, Match_folder_list):
    Version = "11.0.11.1111"  # Default version
    logging.debug('CheckMEVersion Start.')
    for Fv in Match_folder_list:
        path = "..\\" if not os.path.isdir(".\\" + Fv) else ".\\"
        if Fv.split("_")[1] == NProc[2]:
            if os.path.isdir(path + Fv + "\\ME") and os.path.isfile(path + Fv + "\\ME\\ME.bin"):
                MEbinary_pattern = r"ME_+[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]+.bin"
                for root, dirs, files in os.walk(path + Fv + "\\ME"):
                    for name in files:
                        searchObj = search(MEbinary_pattern, name)
                        if searchObj:
                            ME_Version = name.split(".bin")[0]  # Remove .bin
                            Version_List = ME_Version.split("_")
                            for Version in Version_List:
                                if search(r"[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]", Version):
                                    return Version
            elif Version == "11.0.11.1111" and os.path.isfile(path + Fv + "\\ME\\ME.inf"):
                pattern = r"DriverVer.+"
                with open(path + Fv + "\\ME\\ME.inf", "r+") as file:
                    Str_list = file.read()
                    searchObj = search(pattern, Str_list)
                    if searchObj:
                        Version = searchObj.group(0)[25:27] + ".0." + searchObj.group(0)[28:30] + "." + searchObj.group(0)[31:33] + searchObj.group(0)[34:36]
                return Version
            else:
                return Version

# 修改 Release Note
def ModifyReleaseNote(NProc, ReleaseFileName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list):
    print("Platform ReleaseNote Modify...")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    filepath = r"{}".format(ReleaseFileName)
    logging.debug('app books open...')
    logging.debug("Platform_Flag = " + str(Platform_Flag(NProc)))
    print("Please wait for a moment while it is being processed...")
    wb = app.books.open(filepath)
    logging.debug('app books open finish.')

    try:
        # For 1.0 Release note
        if "v1.08" in wb.sheets['Revision'].range('A8').value or \
           "v1.07" in wb.sheets['Revision'].range('A8').value or \
           "v1.06" in wb.sheets['Revision'].range('A8').value:
            handle_release_note_1_0(wb, NProc, ReleaseFileName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
        # For 2.0 Release note
        elif "v2.0" in wb.sheets['Revision'].range('A8').value:
            handle_release_note_2_0(wb, NProc, ReleaseFileName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
        else:
            Old_ReleaseNoteVersion = wb.sheets['Revision'].range('A8').value
            print("ReleaseNote Version is " + Old_ReleaseNoteVersion + " Not in 1.06~1.08 Version, Modify Skip.\n")

        wb.save()
        print("Platform ReleaseNote Modify " + Fore.GREEN + "succeeded.\n")
    except Exception as e:
        logging.error(f"Failed to modify release note: {e}")
        print("Platform ReleaseNote Modify " + Fore.RED + "Failed!\n")
    finally:
        wb.close()
        app.quit()

def handle_release_note_1_0(wb, NProc, ReleaseFileName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list):
    if Platform_Flag(ReleaseFileName) in Intel_Platforms_G5later:
        IntelHistory, IntelHowToFlash = get_sheets(wb, ['IntelPlatformHistory_FY24', 'IntelPlatformHowToFlash_FY24'])
        IntelHistory = IntelHistory or find_sheet_by_partial_name(wb, 'IntelPlatformHistory')
        IntelHowToFlash = IntelHowToFlash or find_sheet_by_partial_name(wb, 'IntelPlatformHowToFlash')

        copy_values_to_new_column(IntelHistory, 'B9:B100', 'C9:C100')
        modify_system_bios_version(IntelHistory, wb.sheets['IntelProjectPN'], NewVersion, NewBuildID)
        Modify_Excel_Data(IntelHistory, 'Build Date', BiosBuildDate[ReleaseFileName.split("_")[2]])
        Modify_Excel_Data(IntelHistory, 'CHECKSUM', "0x" + BiosBinaryChecksum[NProc[2]].upper())
        Modify_Excel_Data(IntelHistory, 'MRC', BiosMrcVersion)
        Modify_Excel_Data(IntelHistory, 'ISH FW version', "HpSigned_ishC_" + BiosIshVersion if Platform_Flag(ReleaseFileName) in Intel_Platforms_G9later else "" + BiosIshVersion)
        Modify_Excel_Data(IntelHistory, 'PMC', BiosPmcVersion)
        Modify_Excel_Data(IntelHistory, 'NPHY', BiosNphyVersion)
        Modify_Excel_Data(wb.sheets['IntelProjectPN'], 'PART NUMBER', "BIOS P00000-000")
        Modify_Excel_Data(wb.sheets['IntelProjectPN'], 'ID', "BIOS 000000")
        handle_me_version(wb.sheets['IntelProjectPN'], CheckMEVersion(NProc, Match_folder_list))
        handle_folder_path(IntelHistory, NProc, ReleaseFileName)
        handle_how_to_flash(IntelHowToFlash, NewVersion, NewBuildID)
        handle_how_to_flash_picture_position(IntelHowToFlash)
    elif Platform_Flag(ReleaseFileName) in AMD_Platforms_R24later:
        AMDHistory, AMDHowToFlash = get_sheets(wb, ['AMDPlatformHistory', 'AMDPlatformHowToFlash'])
        copy_values_to_new_column(AMDHistory, 'B9:B100', 'C9:C100')
        modify_system_bios_version(AMDHistory, None, NewVersion, NewBuildID)
        Modify_Excel_Data(AMDHistory, 'Build Date', BiosBuildDate[ReleaseFileName.split("_")[2]])
        try:
            Modify_Excel_Data(AMDHistory, 'CHECKSUM', "0x" + BiosBinaryChecksum[NProc[2]].upper())
        except KeyError:
            Modify_Excel_Data(AMDHistory, 'CHECKSUM', "0x")
        handle_folder_path(wb.sheets['AMDPlatformInfo'], NProc, ReleaseFileName)
        handle_how_to_flash(AMDHowToFlash, NewVersion, NewBuildID)
        handle_how_to_flash_picture_position(AMDHowToFlash)

def handle_release_note_2_0(wb, NProc, ReleaseFileName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list):
    if Platform_Flag(ReleaseFileName) in Intel_Platforms_G9later:
        ComponentInfo, PlatformInfo, PlatformHistory, PlatformHowToFlash = get_sheets(wb, ['ComponentInfo', 'PlatformInfo', 'PlatformHistory', 'PlatformHowToFlash'])

        copy_values_to_new_column(PlatformHistory, 'B9:B200', 'C9:C200')
        modify_system_bios_version(PlatformHistory, ComponentInfo, NewVersion, NewBuildID)
        Modify_Excel_Data(PlatformHistory, 'Build Date', BiosBuildDate)
        Modify_Excel_Data(PlatformHistory, 'CHECKSUM', "0x" + BiosBinaryChecksum[NProc[2]].upper())
        Modify_Excel_Data(PlatformHistory, 'MRC', BiosMrcVersion)
        Modify_Excel_Data(PlatformHistory, 'ISH FW version', "HpSigned_ishC_" + BiosIshVersion)
        Modify_Excel_Data(PlatformHistory, 'PMC', BiosPmcVersion)
        Modify_Excel_Data(PlatformHistory, 'NPHY FW version', BiosNphyVersion)
        Modify_Excel_Data(PlatformHistory, 'System BIOS', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + ".00")
        Modify_Excel_Data(PlatformHistory, 'HP System Firmware', NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + ".00")
        Modify_Excel_Data(ComponentInfo, 'PART NUMBER', "BIOS P00000-000")
        Modify_Excel_Data(ComponentInfo, 'ID', "BIOS 000000")
        handle_me_version(ComponentInfo, CheckMEVersion(NProc, Match_folder_list))
        handle_folder_path(PlatformInfo, NProc, ReleaseFileName)
        handle_how_to_flash(PlatformHowToFlash, NewVersion, NewBuildID)
        handle_how_to_flash_picture_position(PlatformHowToFlash)

def get_sheets(wb, sheet_names):
    sheets = {name: None for name in sheet_names}
    for sheet in wb.sheets:
        if sheet.name in sheet_names:
            sheets[sheet.name] = sheet
            logging.debug(f'Found sheet: {sheet.name}')
    return sheets[sheet_names[0]], sheets[sheet_names[1]]

def find_sheet_by_partial_name(wb, partial_name):
    for sheet in wb.sheets:
        if partial_name in sheet.name:
            logging.debug(f'Found sheet: {sheet.name}')
            return sheet
    return None

def copy_values_to_new_column(sheet, old_range, new_range):
    sheet.range('C:C').api.Insert(constants.InsertShiftDirection.xlShiftToRight)
    copy_values = sheet.range(old_range).options(ndim=2).value
    sheet.range(new_range).value = copy_values
    sheet.range(new_range).api.Font.Color = 0x000000
    logging.debug('Copied values to new column.')

def modify_system_bios_version(history_sheet, project_pn_sheet, new_version, new_build_id):
    check = ""
    if new_build_id == "" or new_build_id == "0000":
        Modify_Excel_Data(history_sheet, 'System BIOS Version', new_version[0:2] + "." + new_version[2:4] + "." + new_version[4:6])
        if project_pn_sheet:
            Modify_Excel_Data(project_pn_sheet, 'VERSION', new_version[0:2] + "." + new_version[2:4] + "." + new_version[4:6])
        Modify_Excel_Data(history_sheet, 'BIOS Build Version', "0000")
    else:
        Modify_Excel_Data(history_sheet, 'System BIOS Version', new_version[0:2] + "." + new_version[2:4] + "." + new_version[4:6] + "_" + new_build_id)
        if project_pn_sheet:
            Modify_Excel_Data(project_pn_sheet, 'VERSION', new_version[0:2] + "." + new_version[2:4] + "." + new_version[4:6] + "_" + new_build_id)
        Modify_Excel_Data(history_sheet, 'BIOS Build Version', new_build_id)
    check = "pass"
    return check

def handle_me_version(sheet, me_version):
    if me_version != "11.0.11.1111":
        pattern = r'[0-9]+[\.][0-9]+[\.][0-9]+[\.]\d{4}'
        old_me_version = Find_Old_ME_Version(sheet)
        logging.debug('ME Version : ' + me_version)
        logging.debug('Old ME Version : ' + old_me_version)
        Modify_Excel_Data(sheet, 'VERSION', me_version)
        if me_version != old_me_version:
            Modify_Excel_Data(sheet, 'PART NUMBER', "ME P00000-000")
            Modify_Excel_Data(sheet, 'ID', "ME 000000")
    else:
        logging.debug("ME Version not updated, default version detected.")

def handle_folder_path(info_sheet, nproc, release_file_name):
    check = ""
    for a in range(20, 40):
        if info_sheet.range('A' + str(a)).value == 'Folder Path' and \
           info_sheet.range('A' + str(a + 1)).value == 'ODM FTP' and \
           info_sheet.range('A' + str(a + 2)).value == 'Folder Path':
            folder_path_hp = info_sheet.range('B' + str(a)).value
            folder_path_quanta = info_sheet.range('B' + str(a + 2)).value
            pattern = r'\w+_\w+_\w\d\d_\w+.7z'
            re_folder_path_hp = sub(pattern, ("_").join(nproc) + ".7z", folder_path_hp)
            re_folder_path_quanta = sub(pattern, ("_").join(nproc) + ".7z", folder_path_quanta)
            info_sheet.range('B' + str(a)).value = re_folder_path_hp
            info_sheet.range('B' + str(a + 2)).value = re_folder_path_quanta
            logging.debug('Folder Path fill finish.')
            check = "pass"
            break
    if check != "pass":
        print("Can't find ['Folder Path', 'ODM FTP', 'Folder Path']")

def handle_how_to_flash(how_to_flash_sheet, new_version, new_build_id):
    check = ""
    for a in range(10, 35):
        if how_to_flash_sheet.range('A' + str(a)).value == "BIOS Flash: From -> To ":
            how_to_flash_sheet.range(str(a + 2) + ':' + str(a + 2)).api.Insert()
            copy_values = how_to_flash_sheet.range('A' + str(a + 1) + ':K' + str(a + 1)).options(ndim=2).value
            how_to_flash_sheet.range('A' + str(a + 2) + ':K' + str(a + 2)).expand('table').value = copy_values[0]
            pattern = r'\d\d.\d\d.\d\d.*'
            if "->" not in str(how_to_flash_sheet.range('A' + str(a + 1)).value):
                how_to_flash_sheet.range('A' + str(a + 1)).value = "00.00.00-> " + str(how_to_flash_sheet.range('A' + str(a + 1)).value)
            flash_version_left = sub(pattern, how_to_flash_sheet.range('A' + str(a + 1)).value.split("->")[1].strip(), how_to_flash_sheet.range('A' + str(a + 1)).value.split("->")[0])
            if new_build_id == "" or new_build_id == "0000":
                flash_version_right = sub(pattern, new_version[0:2] + "." + new_version[2:4] + "." + new_version[4:6], how_to_flash_sheet.range('A' + str(a + 1)).value.split("->")[1])
            else:
                flash_version_right = sub(pattern, new_version[0:2] + "." + new_version[2:4] + "." + new_version[4:6] + "_" + new_build_id, how_to_flash_sheet.range('A' + str(a + 1)).value.split("->")[1])
            how_to_flash_sheet.range('A' + str(a + 1)).value = flash_version_left + "->" + flash_version_right
            logging.debug('HowToFlash fill finish.')
            check = "pass"
            break
    if check != "pass":
        print("Can't find ['BIOS Flash: From -> To ']")

def handle_how_to_flash_picture_position(how_to_flash_sheet):
    flag = "0"
    for a in range(18, 100):
        if how_to_flash_sheet.range('E' + str(a)).value == "The User Account Control (UAC) setting needs to be set to DISABLED.":
            pic_row = str(a - 1)
            flag = "1"
            break
    logging.debug('Find image location.')
    if flag == "1" and len(how_to_flash_sheet.pictures) == 1:
        pic_name = how_to_flash_sheet.pictures[0].name
    else:
        flag = "0"
    logging.debug('Find image name. flag:' + flag)
    if flag == "1":
        for _ in range(10):
            if int(how_to_flash_sheet.pictures[0].top) != int(how_to_flash_sheet.range('E' + pic_row).top + 13.00):
                how_to_flash_sheet.pictures[0].top = how_to_flash_sheet.range('E' + pic_row).top + 13.00
                sleep(0.5)
                how_to_flash_sheet.pictures[0].top = how_to_flash_sheet.range('E' + pic_row).top + 13.00
    else:
        print("Unable to locate.")
    logging.debug('HowToFlash Pic position modify finish.')
