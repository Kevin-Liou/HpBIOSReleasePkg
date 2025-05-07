#=================================
# coding=UTF-8
# Make Release Pkg Script Main Library.
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com
#=================================
import os
import sys
import glob
import time
import logging
from shutil import copy, copytree, move, rmtree, copy2
from re import sub, search, match, compile

from ReleasePkgLib import *
from .Platform import *
from .Excel import CheckMEVersion


# Exit program.
def ExitProgram(message):
    print(message)
    input("Press any key to exit.")
    sys.exit()


# Working with multiple folders.
# Function to handle file organization across multiple folders.
def MatchMultipleFolder(Match_folder_list, ProcessProjectList, NewVersion):
    print("Check if the folder is multi-layered?")
    TestSignVersion = "84" + NewVersion[-4:] # Test sign binary.
    # Iterate over each folder in the provided folder list.
    for i, Fv_folder in enumerate(Match_folder_list):
        # Walk through the directory structure starting from the current folder.
        for root, dirs, files in os.walk(".\\" + Fv_folder):
            # Iterate over each file in the directories.
            for file in files:
                # Check if the file is binary file. ex: U21_000000_32.bin
                expected_file = f"{ProcessProjectList[i]}_{NewVersion}_32.bin"
                test_sign_file = f"{ProcessProjectList[i]}_{TestSignVersion}_32.bin"
                current_path = os.path.join(root, file)
                #logging.debug(f'expected_file: {expected_file}, test_sign_file: {test_sign_file}')

                # For Production sign binary.
                if file == expected_file:
                    expected_path = os.path.join(".\\", Fv_folder, expected_file)

                    # If the file is in the expected directory, no action is needed.
                    if current_path == expected_path:
                        print("No, not need move files.")
                        continue
                    else:
                        # If the file is not in the expected directory, it needs to be moved.
                        print("Yes, need move files.")
                        # Move all directories and files from the current root to the target directory.
                        target_dir = os.path.join(".\\", Fv_folder)
                        for folder in dirs:
                            move(os.path.join(root, folder), target_dir)
                        for file in files:
                            move(os.path.join(root, file), target_dir)

                        # Check if binary file is now in the correct location.
                        if os.path.exists(expected_path):
                            # Delete the empty source directory if it's different from the target directory
                            if os.path.normpath(root) != os.path.normpath(target_dir):
                                rmtree(root)
                            continue
                        else:
                            # If the file is still not found in the target location, print an error message and exit the program.
                            ExitProgram(f"Can't find {expected_file} in {Fv_folder} folder, Please check Fv folder.")

                # For Test sign binary.
                if file == test_sign_file:
                    expected_path = os.path.join(".\\", Fv_folder, test_sign_file)

                    # If the file is in the expected directory, no action is needed.
                    if current_path == expected_path:
                        print("No, not need move files (Test sign binary).")
                        continue
                    else:
                        # If the file is not in the expected directory, it needs to be moved.
                        print("Yes, need move files (Test sign binary).")
                        # Move test sign binary file from the current root to the target directory.
                        target_dir = os.path.join(".\\", Fv_folder)
                        move(current_path, target_dir)

                        # Check if binary file is now in the correct location.
                        if os.path.exists(expected_path):
                            # Delete the empty source directory if it's different from the target directory
                            rmtree(root)
                            continue
                        else:
                            # If the file is still not found in the target location, print an error message and exit the program.
                            ExitProgram(f"Can't find {test_sign_file} in {Fv_folder} folder, Please check Fv folder.")


# Check Bios Version.
def CheckBiosVersion(OldVersion, NewVersion, NewBuildID, ProcessProject):
    flag = True
    version_pattern = compile(r'^\d{4}$|^\d{6}$')
    if not match(version_pattern, OldVersion):
        print(f"OldVersion {OldVersion} should be 4 or 6 digits.")
        flag = False
    if not match(version_pattern, NewVersion):
        print(f"NewVersion {NewVersion} should be 4 or 6 digits.")
        flag = False

    build_id_pattern = compile(r'^\d{4}$|^$')
    if not match(build_id_pattern, NewBuildID):
        print(f"NewBuildID {NewBuildID} should be 4 digits or empty.")
        flag = False

    project_pattern = compile(r'^[a-zA-Z]\d{2}$')
    if not match(project_pattern, ProcessProject):
        print(f"ProcessProject {ProcessProject} should be 1 letter followed by 2 digits.")
        flag = False

    if flag == False:
        ExitProgram("")


# Safe rename file.
def SafeRename(src, dst):
    # Attempt to rename the file and retry later if it fails.
    max_attempts = 5
    for attempt in range(max_attempts):
        try:
            os.rename(src, dst)
            return
        except PermissionError:
            if attempt < max_attempts - 1:
                time.sleep(1)  # wait 1 second before retrying
            else:
                raise

# Modify Update Version message in file.
def ChangeBuildID(NewProcPkgInfo, Version_file_list, NewVersion):
    pattern = r'\w\d{2}_\d{6}' # For AMD
    PlatID = NewProcPkgInfo[0]
    if Platform_Flag(NewProcPkgInfo) == "R24":
        pattern = r'\w\d{2}_\d{6}'
        PlatID = NewProcPkgInfo[1]
    if (Platform_Flag(NewProcPkgInfo) in Intel_Platforms_G4later) or (Platform_Flag(NewProcPkgInfo) in AMD_Platforms_G12later):
        pattern = r'\w\d{2}_\d{6}'
        PlatID = NewProcPkgInfo[2]
    if Platform_Flag(NewProcPkgInfo) == "Intel G3":
        pattern = r'\w\d{2}_\d{4}'
        PlatID = NewProcPkgInfo[2]
    for Filename in Version_file_list:
        if os.path.isfile(Filename):
            File = open(Filename,"r+")
            Str_list = File.read()
            ReStr_list = sub(pattern, PlatID + "_" + NewVersion, Str_list)
            if ReStr_list != Str_list:
                File.seek(0, 0)
                File.write(ReStr_list)
                File.close()
            else:
                print("Now Pkg Version is already " + str(NewVersion) + ".")
                File.close()
                break
    print("Modify " + ("_").join(NewProcPkgInfo) + " Pkg Version (.nsh & .bat) to " + str(NewVersion) + " succeeded.\n")


# Remove old file in pkg.
def RemoveOldFileInDir(target_dir, RemoveRule, NotRemoveRule):
    i = 0
    for root,dirs,files in os.walk(target_dir):
        for name in files: # Here are the rules for remove
            Path = ".\\" + os.path.join(root, name)
            for Rule_remove in RemoveRule:
                pattern_remove = Rule_remove
                searchObj_remove = search(pattern_remove, Path)
                if os.path.isfile(Path) and searchObj_remove != None: # If Rule_remove in name
                    for Rule_not_remove in NotRemoveRule:
                        pattern_not_remove = Rule_not_remove
                        searchObj_not_remove = search(pattern_not_remove, Path)
                        if searchObj_not_remove != None: # If searchObj_not_remove != None , Break, Stop comparing and remove
                            break
                    if searchObj_not_remove == None:
                        os.remove(Path)
                        i = i + 1
                        print(os.path.join(root, name) + "\t remove succeeded.")
    if i == 0:
        print(target_dir + "\t no file can remove.")
    print()


# Copy old version folder to new version folder.
def Copy_Release_Folder(sourcePath, targetPath):
    print("Start Copy " + sourcePath.split("\\")[-1] + " to " + targetPath.split("\\")[-1] + ", Please wait.....")

    if not os.path.exists(targetPath):
        os.makedirs(targetPath)

    # Copy to new Pkg
    for item in os.listdir(sourcePath):
        logging.debug(f'Copy item: {item}')
        s = os.path.join(sourcePath, item)
        d = os.path.join(targetPath, item)
        if os.path.isdir(s):
            if not item.startswith('~$'): # Ignore ~$ file, ex:~$Arya_PV_S11_BIOS_Release_Note.xlsm
                copytree(s, d, dirs_exist_ok=True)
        else:
            if not item.startswith('~$'): # Ignore ~$ file, ex:~$Arya_PV_S11_BIOS_Release_Note.xlsm
                copy2(s, d)

    print("Copy Pkg " + sourcePath.split("\\")[-1] + " to " + targetPath.split("\\")[-1] + " succeeded.\n")


# Not use now.
# If Fv folder is new folder.
# def New_FvFolder_Move_File(Fv_Path):
#     if os.path.isdir(Fv_Path + "\\Combined\\WU") and os.path.isfile(Fv_Path + "\\Combined\\WU\\fwu.pvk"):
#         for root, dirs, files in os.walk(Fv_Path + "\\Combined\\WU"):
#             for name in files: # Move "WU" file.
#                 if os.path.isfile(Fv_Path + "\\Combined\\WU\\" + name):
#                     print("move " + Fv_Path + "\\Combined\\WU\\" + name + " to " + Fv_Path + "\\Combined")
#                     move(Fv_Path + "\\Combined\\WU\\" + name, Fv_Path + "\\Combined")
#                 elif os.path.isfile(Fv_Path + "Combined\\WU\\" + name):
#                     print("move " + Fv_Path + "Combined\\WU\\" + name + " to " + Fv_Path + "\\Combined")
#                     move(Fv_Path + "Combined\\WU\\" + name, Fv_Path + "\\Combined")


# Modify ME version in WLAN_MCU files.
def Modify_ME_WLAN_MCU_Files(path, ME_Version):
    files_to_check = ['WLAN_MCU.bat', 'WLAN_MCU.nsh', 'WLAN_MCU64.bat']
    pattern = compile(r'ME_\d+\.\d+\.\d+\.\d+\.bin')

    for file_name in files_to_check:
        file_path = os.path.join(path, file_name)
        if os.path.isfile(file_path):
            with open(file_path, 'r') as file:
                content = file.read()

            new_content = pattern.sub(f'ME_{ME_Version}.bin', content)

            if new_content != content:
                with open(file_path, 'w') as file:
                    file.write(new_content)
                print(f"Updated {file_name} with ME version {ME_Version}")


# Copy Fv folder file to NewPkg.
def Copy_Release_Files(sourceFolder, targetFolder, NProc, Match_folder_list):
    source_fullpath = ".\\" + sourceFolder + "\\"
    target_fullpath = ".\\" + targetFolder + "\\"
    logging.debug(f'source_fullpath: {source_fullpath}, target_fullpath: {target_fullpath}')
    # Combined copy to Capsule&HPFWUPDREC.
    #======For G5 and late Fv
    # Check Capsule folder format for G5 and late.
    if Platform_Flag(targetFolder) in Intel_Platforms_G5later:
        if not os.path.isdir(target_fullpath + "\\Capsule\\Linux"):
            os.makedirs(target_fullpath + "\\Capsule\\Linux")
            os.makedirs(target_fullpath + "\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)")
        # Check Capsule folder format for G6 and late.
        if os.path.isdir(target_fullpath + "\\Capsule\\CCG5") and Platform_Flag(targetFolder) in Intel_Platforms_G6later:
            os.rename(target_fullpath + "\\Capsule\\CCG5", target_fullpath + "\\Capsule\\PD_FW")
        if not os.path.isdir(target_fullpath + "\\Capsule\\Windows"):
            os.makedirs(target_fullpath + "\\Capsule\\Windows")
            os.makedirs(target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
            os.makedirs(target_fullpath + "\\Capsule\\Windows\\Thunderbolt")
        for file in glob.glob(target_fullpath + "\\Capsule\*.doc*"):
            if "submission" in file or "Submission" in file:
                move(file, target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
                print("move " + file + " to " + target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")

        # Copy FUR and WU files.
        if os.path.isdir(source_fullpath + "\\Combined\\FUR") and os.path.isdir(source_fullpath + "\\Combined\\WU"):
            for root,dirs,files in os.walk(source_fullpath + "\\Combined\\FUR"):
                for name in files:
                    if ".bin" in name or ".inf" in name:
                        copy(root + "\\" + name, target_fullpath + "\\HPFWUPDREC")
                        print(root + "\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
            for root,dirs,files in os.walk(source_fullpath + "\\Combined\\WU"):
                for name in files:
                    copy(root + "\\" + name, target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)" + " Copy succeeded.")

        # If Linux folder exist, copy files.
        if os.path.isdir(source_fullpath + "\\Combined\\Linux"):
            for root,dirs,files in os.walk(source_fullpath + "\\Combined\\Linux"):
                for name in files:
                    copy(root + "\\" + name, target_fullpath + "\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)" + " Copy succeeded.")

        # If TBT FW exist, copy it, for Intel G5 and late if support TBT use.
        TBT_path = ""
        TBT_Version = ""
        if not os.path.exists(target_fullpath + "\\Capsule\\Windows\\Thunderbolt"):
            os.makedirs(target_fullpath + "\\Capsule\\Windows\\Thunderbolt")
        if os.path.isdir(source_fullpath + "\\TBT"):
            TBT_path = source_fullpath + "\\TBT"
        elif os.path.isdir(source_fullpath + "\\TBT_RTD3"):
            TBT_path = source_fullpath + "\\TBT_RTD3"
        if not TBT_path == "":
            for root,dirs,files in os.walk(TBT_path):
                for name in files:
                    if os.path.isfile(target_fullpath + "\\Capsule\\Windows\\Thunderbolt\\" + name):
                        os.remove(target_fullpath + "\\Capsule\\Windows\\Thunderbolt\\" + name)
                    copy(root + "\\" + name, target_fullpath + "\\Capsule\\Windows\\Thunderbolt")
                    print(root + "\\" + name + " to " + target_fullpath + "\\Capsule\\Windows\\Thunderbolt" + " Copy succeeded.")
                    if ".inf" in name:
                        Pattern = r"HP_\w+_\w+_\w+_\w+_\w+_\w+_\d+"
                        File = open(root + "\\" + name,"r+")
                        Str_list = File.read()
                        searchObj = search(Pattern, Str_list)
                        if searchObj != None:
                            TBT_Version = searchObj.group(0)
                            File.close()
            if TBT_Version != "":
                Dirs = os.listdir(target_fullpath + "\\Capsule\\Windows\\Thunderbolt")
                for file in Dirs:
                    if ".bin" in file:
                        os.rename(target_fullpath + "\\Capsule\\Windows\\Thunderbolt\\" + file, target_fullpath + "\\Capsule\\Windows\\Thunderbolt\\" + TBT_Version + ".bin")
                    elif ".inf" in file:
                        os.rename(target_fullpath + "\\Capsule\\Windows\\Thunderbolt\\" + file, target_fullpath + "\\Capsule\\Windows\\Thunderbolt\\" + TBT_Version + ".inf")
            if os.path.exists(target_fullpath + "\\Capsule\\TBT"):
                rmtree(target_fullpath + "\\Capsule\\TBT")
        else:
            print("Can't find TBT folder.")

        # ME binary copy to METools\FWUpdate\HPSignME for Intel G5 and late.
        MEbinary_pattern = r"ME_+[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]+.bin"
        if not os.path.exists(target_fullpath + "\\METools\\FWUpdate\\HPSignME"): # Copy sign ME file
            os.makedirs(target_fullpath + "\\METools\\FWUpdate\\HPSignME")
            logging.debug('Make dirs \\METools\\FWUpdate\\HPSignME.')
        ME_Bin_Check = "False"
        for root,dirs,files in os.walk(source_fullpath + "\\ME"):
            for name in files:
                searchObj = search(MEbinary_pattern, name) # Search ME binary file
                if (searchObj != None):
                    copy(root + "\\" + name, target_fullpath + "\\METools\\FWUpdate\\HPSignME")
                    print(root + "\\" + name + "(Sign) to " + targetFolder + "\\METools\\FWUpdate\\HPSignME" + " Copy succeeded.")
                    ME_Version = CheckMEVersion(NProc, Match_folder_list) # ex. 14.0.21.7227
                    ME_Bin_Check = "True"
                    logging.debug('ME_Version:' + ME_Version)
                    break
        # Copy sign ME file
        if ME_Bin_Check == "False" and os.path.isfile(source_fullpath + "\\ME\\ME.bin"):
            ME_filename = "ME.bin"
            if os.path.isfile(source_fullpath + "\\ME\\ME.inf") and os.path.isfile(source_fullpath + "\\ME\\ME.bin"):
                ME_Version = CheckMEVersion(NProc, Match_folder_list) # ex. 14.0.21.7227
                logging.debug('ME_Version:' + ME_Version)
                if not os.path.isfile(source_fullpath + "\\ME\\ME_" + ME_Version + ".bin"):
                    os.rename(source_fullpath + "\\ME\\ME.bin", source_fullpath + "\\ME\\ME_" + ME_Version + ".bin")
                    print("Rename: ME.bin to ME_" + ME_Version + ".bin")
                    ME_filename = "ME_" + ME_Version + ".bin"
            copy(source_fullpath + "\\ME\\" + ME_filename, target_fullpath + "\\METools\\FWUpdate\\HPSignME")
            print(root + "\\" + ME_filename + "(Sign) to " + targetFolder + "\\METools\\FWUpdate\\HPSignME" + " Copy succeeded.")
        # Copy unsign ME file
        if os.path.isfile(source_fullpath + "\\ME\\ME_0101.bin"):
            if os.path.isfile(target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin"):
                os.remove(target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin")
            copy(source_fullpath + "\\ME\\ME_0101.bin", target_fullpath + "\\METools\\FWUpdate\\MEFW")
            os.rename(target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_0101.bin", target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin")
            print(sourceFolder + "\\ME\\" + ME_Version + "(UnSign) to " + targetFolder + "\\METools\\FWUpdate\\MEFW" + " Copy succeeded.")
            # Modify the contents of the WLAN_MCU files.
            if os.path.isfile(target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin"):
                MEFW_path = target_fullpath + "\\METools\\FWUpdate\\MEFW\\"
                Modify_ME_WLAN_MCU_Files(MEFW_path, ME_Version)

    #=======For G4 Fv
    if Platform_Flag(targetFolder) == "Intel G4":
        if os.path.isfile(source_fullpath + "\\Combined\\fwu.pfx"):
            for root,dirs,files in os.walk(source_fullpath + "\\Combined"):
                for name in files:
                    if ".bin" in name or ".inf" in name:
                        copy(root + "\\" + name, target_fullpath + "\\HPFWUPDREC")
                        print(root + "\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
                    copy(root + "\\" + name, target_fullpath  + "\\Capsule")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule" + " Copy succeeded.")

    #======For G3 FV
    if Platform_Flag(targetFolder) == "Intel G3":
        if os.path.isfile(source_fullpath + "\\fwu.pfx"):
            for root,dirs,files in os.walk(source_fullpath):
                for name in files:
                    if "_12.bin" in name:
                        copy(root + "\\" + name, target_fullpath + "\\HPBIOSUPDREC")
                        os.rename(target_fullpath + "\\HPBIOSUPDREC\\" + name, target_fullpath + "\\HPBIOSUPDREC\\" + name[0:8] + ".bin")
                        print(sourceFolder + "\\" + name[0:8] + ".bin" + " to " + targetFolder + "\\HPBIOSUPDREC" + " Copy succeeded.")
                        copy(root + "\\" + name, target_fullpath + "\\Capsule Update")
                        os.rename(target_fullpath + "\\Capsule Update\\" + name, target_fullpath + "\\Capsule Update\\" + name[0:8] + ".bin")
                        print(sourceFolder + "\\" + name[0:8] + ".bin" + " to " + targetFolder + "\\Capsule Update" + " Copy succeeded.")
                    if ".inf" in name:
                        copy(root + "\\" + name, target_fullpath + "\\HPBIOSUPDREC")
                        print(sourceFolder + "\\" + name + " to " + targetFolder + "\\HPBIOSUPDREC" + " Copy succeeded.")
                    if ".cer" in name or ".pfx" in name or ".pvk" in name or ".cat" in name or ".inf" in name:
                        copy(root + "\\" + name, target_fullpath + "\\Capsule Update")
                        print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Capsule Update" + " Copy succeeded.")

    # Bin file copy to FPTW&Global.
    for name in os.listdir(source_fullpath):
        source_file = os.path.join(source_fullpath, name)

        # Check if the file needs to be copied to the FPTW folder, _84 files test sign binary are excluded.
        if "_84" not in name and ("_32.bin" in name or "_12.bin" in name or "_16.bin" in name):
            target_dir = os.path.join(target_fullpath, "FPTW")
            copy(source_file, target_dir)
            print(f"{source_file} to {target_dir} Copy succeeded.")

        # Check if the file needs to be copied to the Global folder, _84 files test sign binary are excluded.
        if "_84" not in name and ("_12.bin" in name or "_32.bin" in name):
            target_dir = os.path.join(target_fullpath, "Global", "BIOS")
            copy(source_file, target_dir)
            print(f"{source_file} to {target_dir} Copy succeeded.")

        # Check if the file needs to be copied to the XML folder
        if ".xml" in name and str(source_fullpath.split("_")[1]) in name:
            target_dir = os.path.join(target_fullpath, "XML")
            copy(source_file, target_dir)
            print(f"{source_file} to {target_dir} Copy succeeded.")

        # Check if the file is a Pvt. bin file and if it needs to be copied to the root directory
        if "Pvt.bin" in name:
            target_dir = target_fullpath
            copy(source_file, target_dir)
            print(f"{source_file} to {target_dir} Copy succeeded.")

        # Check if the file is a test signature binary
        if "_84" in name and "_32.bin" in name and os.path.exists(os.path.join(target_fullpath, "TestSign")):
            target_dir = os.path.join(target_fullpath, "TestSign")
            copy(source_file, target_dir)
            print(f"{source_file} to {target_dir} Copy succeeded.")
    print("Copy process completed.\n")


# Copy Fv folder file to NewPkg.(For AMD)
def Copy_Release_Files_AMD(sourceFolder, targetFolder, NewVersion):
    source_fullpath = ".\\" + sourceFolder + "\\"
    target_fullpath = ".\\" + targetFolder + "\\"
    # Combined copy to Capsule&HPFWUPDREC.
    #======For G5 and late Fv
    if os.path.isdir(source_fullpath + "\\Combined\\FUR") and os.path.isdir(source_fullpath + "\\Combined\\WU"):
        for root,dirs,files in os.walk(source_fullpath + "\\Combined\\FUR"):
            for name in files:
                if ".bin" in name or ".inf" in name:
                    copy(source_fullpath + "Combined\\FUR\\" + name, target_fullpath + "\\HPFWUPDREC")
                    print(sourceFolder + "\\Combined\\FUR\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
        for root,dirs,files in os.walk(source_fullpath + "\\Combined\\WU"):
            for name in files:
                if Platform_Flag(targetFolder) not in AMD_Platforms_G12later:
                    copy(source_fullpath + "\\Combined\\WU\\" + name, target_fullpath + "\\Capsule\\Windows")
                    print(sourceFolder + "\\Combined\\WU\\" + name + " to " + targetFolder + "\\Capsule\\Windows" + " Copy succeeded.")
                else:
                    copy(source_fullpath + "\\Combined\\WU\\" + name, target_fullpath + "\\Capsule\\Windows\Combined FW Image (BIOS, PD, RETIMER)")
                    print(sourceFolder + "\\Combined\\WU\\" + name + " to " + targetFolder + "\\Capsule\\Windows\Combined FW Image (BIOS, PD, RETIMER)" + " Copy succeeded.")                
    # If Linux folder exist, copy files.
    if os.path.isdir(source_fullpath+"\\Combined\\Linux"):
        for root,dirs,files in os.walk(source_fullpath+"\\Combined\\Linux"):
            for name in files:
                if Platform_Flag(targetFolder) not in AMD_Platforms_G12later:
                    copy(source_fullpath + "\\Combined\\Linux\\" + name, target_fullpath + "\\Capsule\\Linux")
                    print(sourceFolder + "\\Combined\\Linux\\" + name + " to " + targetFolder + "\\Capsule\\Linux" + " Copy succeeded.")
                else:
                    copy(source_fullpath + "\\Combined\\Linux\\" + name, target_fullpath + "\\Capsule\\Linux\Combined FW Image (BIOS, PD, RETIMER)")
                    print(sourceFolder + "\\Combined\\Linux\\" + name + " to " + targetFolder + "\\Capsule\\Linux\Combined FW Image (BIOS, PD, RETIMER)" + " Copy succeeded.")             
    #=======For G4 other Fv
    elif os.path.isfile(source_fullpath + "\\Combined\\fwu.pfx"):
        for root,dirs,files in os.walk(source_fullpath + "\\Combined"):
            for name in files:
                if ".bin" in name or ".inf" in name:
                    copy(source_fullpath + "\\Combined\\" + name, target_fullpath + "\\HPFWUPDREC")
                    print(sourceFolder + "\\Combined\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
                copy(source_fullpath + "\\Combined\\" + name, target_fullpath + "\\Capsule")
                print(sourceFolder + "\\Combined\\" + name + " to " + targetFolder + "\\Capsule" + " Copy succeeded.")
    # Bin file copy to FPTW&Global.
    for name in os.listdir(source_fullpath):# Bin file copy to FPTW&Global
        if Platform_Flag(name) in AMD_Platforms: # 78 AMD 78 R26 78787878787878
            if Platform_Flag(targetFolder) not in AMD_Platforms_G12later:
                if NewVersion + "_16.bin" in name:
                    copy(source_fullpath + name, target_fullpath + "\\AMDFLASH")
                    print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" +" Copy succeeded.")
            if NewVersion + "_32.bin" in name:
                copy(source_fullpath + name, target_fullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" +" Copy succeeded.")
        else:
            if NewVersion + ".bin" in name:
                copy(source_fullpath + name, target_fullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" + " Copy succeeded.")
        if Platform_Flag(targetFolder) not in AMD_Platforms_G12later:
            if NewVersion + ".bin" in name:
                copy(source_fullpath + name, target_fullpath + "\\Global\\BIOS")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.")
        else:
            if NewVersion + "_32.bin" in name:
                copy(source_fullpath + name, target_fullpath + "\\Global\\BIOS")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.") 
        # Copy to XML
        if ".xml" in name:
            if Platform_Flag(targetFolder) not in AMD_Platforms_G12later:
                if str(Platform_Flag(name)) in name:
                    copy(source_fullpath + name, target_fullpath + "\\XML")
                    print(sourceFolder + "\\" + name + " to " + targetFolder + "\\XML" + " Copy succeeded.")
            else:
                if NewVersion + ".xml" in name:
                    copy(source_fullpath + name, target_fullpath + "\\XML")
                    print(sourceFolder + "\\" + name + " to " + targetFolder + "\\XML" + " Copy succeeded.")
        # For Smart flash copy *Pvt.bin.
        if "Pvt.bin" in name:
            copy(source_fullpath + name, target_fullpath)
            print(sourceFolder + "\\" + name + " to " + targetFolder + " Copy succeeded.")
        # Check if the file is a test signature binary
        if "_84" in name and "_32.bin" in name and os.path.exists(os.path.join(target_fullpath, "TestSign")):
            copy(source_fullpath + name, target_fullpath + "\\TestSign")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\TestSign" + " Copy succeeded.")
    print("Copy process completed.\n")


# Find Fv Folder, Add it to the Match_list and return it
def FindFvFolder(ProcessProjectList, NewVersion, NewBuildID):
    Match_list = []
    if len(NewVersion) >= 6:
        Github_PkgVersion = f".{int(NewVersion[0:2])}.{int(NewVersion[2:4])}.{int(NewVersion[4:6])}"

    for i, Process in enumerate(ProcessProjectList):
        for Dir in os.listdir(".\\"):
            Match = f"Fv_{Process}_{NewVersion}_{NewBuildID}".lower() if NewBuildID else f"Fv_{Process}_{NewVersion}_0000".lower()
            Dir_lower = Dir.lower()

            if Match in Dir_lower and os.path.isdir(os.path.join(".\\", Dir)):
                Match_list.append(Dir)
                break

            elif Github_PkgVersion in Dir_lower and not "fv_" in Dir_lower and os.path.isdir(os.path.join(".\\", Dir)):
                Match_list.append(Dir)
                break

    logging.debug(f'Match_list: {Match_list}')
    return Match_list


# Find Fv Zip file, Add it to the Match_list and return it
def FindFvZip(ProcessProjectList, ProjectNameInfo, NewVersion, NewBuildID):
    Match_list = []
    Selected_file = None
    #======For New Github Fv
    if len(NewVersion) >= 6:
        version_base = f".{int(NewVersion[0:2])}.{int(NewVersion[2:4])}.{int(NewVersion[4:6])}"
        OldGithub_PkgVersion = version_base + ".zip"
        NewGithub_PkgVersion = f"{version_base}.{int(NewBuildID)}.zip" if NewBuildID and NewBuildID != "0000" else "FFFFFF"
        logging.debug(f'OldGithub_PkgVersion: {OldGithub_PkgVersion}, NewGithub_PkgVersion: {NewGithub_PkgVersion}')

    for i, project in enumerate(ProcessProjectList):
        # Construct matching strings based on old and new versions
        OldVersionMatch = f"Fv_{project}_{NewVersion}_".lower()
        NewVersionMatch = f"Fv_{project}_{NewVersion}_{NewBuildID}".lower() if NewBuildID else OldVersionMatch

        for Dir in os.listdir(".\\"):
            Dir_lower = Dir.lower()
            # Check if it is an old version of Fv file
            if (OldVersionMatch in Dir_lower or NewVersionMatch in Dir_lower) and (Dir_lower.endswith(".zip") or Dir_lower.endswith(".7z")):
                logging.debug(f'Normal Dir_lower: {Dir_lower}')
                if Dir_lower.endswith(".7z"):
                    new_name = f"{os.path.splitext(Dir)[0]}.zip"
                    os.rename(f".\\{Dir}", f".\\{new_name}")
                    print(f"Rename {Dir} to {new_name}")
                    Dir = new_name
                # If a matching new version is found, prioritize
                if NewVersionMatch in Dir_lower:
                    Match_list = [Dir]  # Clear previous matches and add new matches
                    break
                elif OldVersionMatch in Dir_lower and not Match_list:
                    Match_list.append(Dir)

            # Checking for Github Fv Files
            if (OldGithub_PkgVersion[:-4] in Dir_lower or NewGithub_PkgVersion[:-4] in Dir_lower) and (Dir_lower.endswith(".zip") or Dir_lower.endswith(".7z") or Dir_lower.endswith(".nupkg")):
                logging.debug(f'Github Dir_lower: {Dir_lower}')

                # Rename .nupkg and .7z to .zip
                if (Dir_lower.endswith(".nupkg") or Dir_lower.endswith(".7z")):
                    new_name = f"{os.path.splitext(Dir)[0]}.zip"
                    os.rename(f".\\{Dir}", f".\\{new_name}")
                    print(f"unpkg Rename {Dir} to {new_name}")
                    Dir = new_name

                # If a matching new version is found, prioritize
                logging.debug(f'NewGithub_PkgVersion: {NewGithub_PkgVersion}, OldGithub_PkgVersion: {OldGithub_PkgVersion}, Dir_lower: {Dir_lower}, Selected_file: {Selected_file}')
                if NewGithub_PkgVersion and NewGithub_PkgVersion in Dir.lower():
                    Selected_file = Dir
                # If a matching old version is found, prioritize
                elif OldGithub_PkgVersion in Dir.lower() and not Selected_file:
                    Selected_file = Dir

        # Check selected file name, if not match, rename it.
        if Selected_file:
            if not Selected_file.startswith(project.lower() + ProjectNameInfo[i].lower()):
                new_name = f"{project.lower()}{ProjectNameInfo[i].lower()}{version_base}"
                new_name = new_name + f".{int(NewBuildID)}.zip" if NewBuildID and NewBuildID != "0000" else new_name + ".zip"
                os.rename(f".\\{Selected_file}", f".\\{new_name}")
                print(f"Mistake Rename {Selected_file} to {new_name}")
                Selected_file = new_name
            Match_list = [Selected_file]

    return Match_list