#=================================
# coding=UTF-8
# Make Release Pkg Script Main Library.
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com
#=================================
import sys, os, glob, logging
from shutil import copy, copytree, move, rmtree
from re import sub, search

from ReleasePkgLib import *
from .Platform import Platform_Flag
from .Excel import CheckMEVersion


# Working with multiple folders.
# Function to handle file organization across multiple folders.
def MatchMultipleFolder(Match_folder_list):
    print("Check is multiple folder?")
    # Iterate over each folder in the provided folder list.
    for Fv in Match_folder_list:
        # Walk through the directory structure starting from the current folder.
        for root, dirs, files in os.walk(".\\" + Fv):
            # Iterate over each file in the directories.
            for file in files:
                # Check if the file is 'AutoGenFlashMap.h'.
                if file == "AutoGenFlashMap.h":
                    # If the file is in the expected directory, no action is needed.
                    if (root + "\\" + file) == (".\\" + Fv + "\\" + "AutoGenFlashMap.h"):
                        print("No, not need move.")
                        return
                    else:
                        # If the file is not in the expected directory, it needs to be moved.
                        print("Yes, need move.")
                        # Move all directories and files from the current root to the target directory.
                        for folder in dirs:
                            move(root + "\\" + folder, ".\\" + Fv + "\\")
                        for file in files:
                            move(root + "\\" + file, ".\\" + Fv + "\\")
                        # Check if 'AutoGenFlashMap.h' is now in the correct location.
                        if os.path.exists(".\\" + Fv + "\\" + "AutoGenFlashMap.h"):
                            return
                        else:
                            # If the file is still not found in the target location, print an error message and exit the program.
                            print("Can't find AutoGenFlashMap.h in Fv folder, Please check Fv folder.")
                            sys.exit()

# Modify Update Version message in file.
def ChangeBuildID(NewProcPkgInfo, Version_file_list, NewVersion):
    pattern = r'\w\d{2}_\d{6}' # For AMD
    PlatID = NewProcPkgInfo[0]
    if (Platform_Flag(NewProcPkgInfo) == "R24"):
        pattern = r'\w\d{2}_\d{6}'
        PlatID = NewProcPkgInfo[1]
    if (Platform_Flag(NewProcPkgInfo) == "Intel G4") or (Platform_Flag(NewProcPkgInfo) == "Intel G5") or \
        (Platform_Flag(NewProcPkgInfo) == "Intel G6") or (Platform_Flag(NewProcPkgInfo) == "Intel G8") or \
        (Platform_Flag(NewProcPkgInfo) == "Intel G9") or (Platform_Flag(NewProcPkgInfo) == "Intel G10"):
        pattern = r'\w\d{2}_\d{6}'
        PlatID = NewProcPkgInfo[2]
    if (Platform_Flag(NewProcPkgInfo) == "Intel G3"):
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
        for name in files:# Here are the rules for remove
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
        print(target_dir+"\t no file can remove.")
    print()


# Copy old version folder to new version folder.
def Copy_Release_Folder(sourcePath, targetPath):
    print("Start Copy "+sourcePath.split("\\")[-1] + " to "+targetPath.split("\\")[-1] + ", Please wait.....")
    copytree(sourcePath, targetPath)# Copy to new Pkg
    print("Copy Pkg " + sourcePath.split("\\")[-1] + " to " + targetPath.split("\\")[-1] + " succeeded.\n")


# Not use now.
# If Fv folder is new folder.
# def New_FvFolder_Move_File(Fv_Path):
#     if os.path.isdir(Fv_Path + "\\Combined\\WU") and os.path.isfile(Fv_Path + "\\Combined\\WU\\fwu.pvk"):
#         for root, dirs, files in os.walk(Fv_Path + "\\Combined\\WU"):
#             for name in files:# Move "WU" file.
#                 if os.path.isfile(Fv_Path + "\\Combined\\WU\\" + name):
#                     print("move " + Fv_Path + "\\Combined\\WU\\" + name + " to " + Fv_Path + "\\Combined")
#                     move(Fv_Path + "\\Combined\\WU\\" + name, Fv_Path + "\\Combined")
#                 elif os.path.isfile(Fv_Path + "Combined\\WU\\" + name):
#                     print("move " + Fv_Path + "Combined\\WU\\" + name + " to " + Fv_Path + "\\Combined")
#                     move(Fv_Path + "Combined\\WU\\" + name, Fv_Path + "\\Combined")


# Copy Fv folder file to NewPkg.
def Copy_Release_Files(sourceFolder, targetFolder, NProc, Match_folder_list):
    source_fullpath = ".\\" + sourceFolder + "\\"
    target_fullpath = ".\\" + targetFolder + "\\"
    # Combined copy to Capsule&HPFWUPDREC.
    #======For G5 and late Fv
    # Check Capsule folder format for G5 and late.
    if (Platform_Flag(targetFolder) == "Intel G5") or (Platform_Flag(targetFolder) == "Intel G6") or\
        (Platform_Flag(targetFolder) == "Intel G8") or (Platform_Flag(targetFolder) == "Intel G9") or\
        (Platform_Flag(targetFolder) == "Intel G10"):
        if os.path.isdir(target_fullpath + "\\Capsule\\CCG5") and ((Platform_Flag(targetFolder) == "Intel G6") or \
            (Platform_Flag(targetFolder) == "Intel G8") or (Platform_Flag(targetFolder) == "Intel G9") or \
            (Platform_Flag(targetFolder) == "Intel G10")):
            os.rename(target_fullpath + "\\Capsule\\CCG5", target_fullpath + "\\Capsule\\PD_FW")
        if not os.path.isdir(target_fullpath + "\\Capsule\\Windows"):
            os.makedirs(target_fullpath + "\\Capsule\\Windows")
            os.makedirs(target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
            os.makedirs(target_fullpath + "\\Capsule\\Windows\\Thunderbolt")
        if not os.path.isdir(target_fullpath + "\\Capsule\\Linux"):
            os.makedirs(target_fullpath + "\\Capsule\\Linux")
            os.makedirs(target_fullpath + "\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)")
        for file in glob.glob(target_fullpath + "\\Capsule\*.doc*"):
            if file.find("submission") != -1 or file.find("Submission") != -1:
                move(file, target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
                print("move " + file + " to " + target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")

        # Copy FUR and WU files.
        if os.path.isdir(source_fullpath + "\\Combined\\FUR") and os.path.isdir(source_fullpath + "\\Combined\\WU"):
            for root,dirs,files in os.walk(source_fullpath + "\\Combined\\FUR"):
                for name in files:
                    if name.find(".bin") != -1 or name.find(".inf") != -1:
                        copy(root + "\\" + name, target_fullpath + "\\HPFWUPDREC")
                        print(root + "\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
            for root,dirs,files in os.walk(source_fullpath + "\\Combined\\WU"):
                for name in files:
                    copy(root + "\\" + name, target_fullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)" + " Copy succeeded.")

        # If Linux folder exist, copy files.
        if os.path.isdir(source_fullpath+"\\Combined\\Linux"):
            for root,dirs,files in os.walk(source_fullpath+"\\Combined\\Linux"):
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

        # ME binary copy to METools\FWUpdate\HPSignME for Intel G5 and late.
        MEbinary_pattern = r"ME_+[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]+.bin"
        if not os.path.exists(target_fullpath + "\\METools\\FWUpdate\\HPSignME"): # Copy sign ME file
            os.makedirs(target_fullpath + "\\METools\\FWUpdate\\HPSignME")
            logging.debug('Make dirs \\METools\\FWUpdate\\HPSignME.')
        ME_Bin_Check = "False"
        for root,dirs,files in os.walk(source_fullpath + "\\ME"):
            for name in files:
                searchObj = search(MEbinary_pattern, name)
                if (searchObj != None):
                    copy(root + "\\" + name, target_fullpath + "\\METools\\FWUpdate\\HPSignME")
                    print(root + "\\" + name + "(Sign) to " + targetFolder + "\\METools\\FWUpdate\\HPSignME" + " Copy succeeded.")
                    ME_Version = CheckMEVersion(NProc, Match_folder_list) # ex. 14.0.21.7227
                    ME_Bin_Check = "True"
                    logging.debug('ME_Version:' + ME_Version)
                    break
        if (ME_Bin_Check == "False") and (os.path.isfile(source_fullpath + "\\ME\\ME.bin")):
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
        if os.path.isfile(source_fullpath + "\\ME\\ME_0101.bin"):# Copy unsign ME file
            if os.path.isfile(target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin"):
                os.remove(target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin")
            copy(source_fullpath + "\\ME\\ME_0101.bin", target_fullpath + "\\METools\\FWUpdate\\MEFW")
            os.rename(target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_0101.bin", target_fullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin")
            print(sourceFolder + "\\ME\\" + ME_Version + "(UnSign) to " + targetFolder + "\\METools\\FWUpdate\\MEFW" + " Copy succeeded.")

    #=======For G4 other Fv
    if (Platform_Flag(targetFolder) == "Intel G4"):
        if os.path.isfile(source_fullpath + "\\Combined\\fwu.pfx"):
            for root,dirs,files in os.walk(source_fullpath + "\\Combined"):
                for name in files:
                    if name.find(".bin") != -1 or name.find(".inf") != -1:
                        copy(root + "\\" + name, target_fullpath + "\\HPFWUPDREC")
                        print(root + "\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
                    copy(root + "\\" + name, target_fullpath  + "\\Capsule")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule" + " Copy succeeded.")

    #======For G3 FV
    if Platform_Flag(targetFolder) == "Intel G3":
        if os.path.isfile(source_fullpath + "\\fwu.pfx"):
            for root,dirs,files in os.walk(source_fullpath):
                for name in files:
                    if name.find("_12.bin") != -1:
                        copy(root + "\\" + name, target_fullpath + "\\HPBIOSUPDREC")
                        os.rename(target_fullpath + "\\HPBIOSUPDREC\\" + name, target_fullpath + "\\HPBIOSUPDREC\\" + name[0:8] + ".bin")
                        print(sourceFolder + "\\" + name[0:8] + ".bin" + " to " + targetFolder + "\\HPBIOSUPDREC" + " Copy succeeded.")
                        copy(root + "\\" + name, target_fullpath + "\\Capsule Update")
                        os.rename(target_fullpath + "\\Capsule Update\\" + name, target_fullpath + "\\Capsule Update\\" + name[0:8] + ".bin")
                        print(sourceFolder + "\\" + name[0:8] + ".bin" + " to " + targetFolder + "\\Capsule Update" + " Copy succeeded.")
                    if name.find(".inf") != -1:
                        copy(root + "\\" + name, target_fullpath + "\\HPBIOSUPDREC")
                        print(sourceFolder + "\\" + name + " to " + targetFolder + "\\HPBIOSUPDREC" + " Copy succeeded.")
                    if name.find(".cer") != -1 or name.find(".pfx") != -1 or name.find(".pvk") != -1 or name.find(".cat" ) != -1 or name.find(".inf") != -1:
                        copy(root + "\\" + name, target_fullpath + "\\Capsule Update")
                        print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Capsule Update" + " Copy succeeded.")

    # Bin file copy to FPTW&Global.
    for name in os.listdir(source_fullpath):
        if name.find("_12.bin") != -1 or name.find("_16.bin") != -1 or name.find("_32.bin") != -1:
            copy(source_fullpath + name, target_fullpath + "\\FPTW")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\FPTW" + " Copy succeeded.")
        if name.find("_12.bin") != -1:# If 16MB BIOS case please add it.
            name = name.split("_12.bin")[0] + "_16.bin"
            copy(source_fullpath + name, target_fullpath + "\\Global\\BIOS")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.")
        elif name.find("_32.bin") != -1:
            copy(source_fullpath + name, target_fullpath + "\\Global\\BIOS")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.")
        if name.find(".xml") != -1:# Copy to XML
            if name.find(str(source_fullpath.split("_")[1])) != -1:
                copy(source_fullpath + name, target_fullpath + "\\XML")
                print(sourceFolder + "\\" + name + " to "+targetFolder + "\\XML" + " Copy succeeded.")
        # For Smart flash copy *Pvt.bin.
        if name.find("Pvt.bin") != -1:
            copy(source_fullpath + name, target_fullpath)
            print(sourceFolder + "\\" + name + " to " + targetFolder + " Copy succeeded.")
    print()


# Copy Fv folder file to NewPkg.(For AMD)
def Copy_Release_Files_AMD(sourceFolder, targetFolder, NewVersion):
    source_fullpath = ".\\" + sourceFolder + "\\"
    target_fullpath = ".\\" + targetFolder + "\\"
    # Combined copy to Capsule&HPFWUPDREC.
    #======For G5 and late Fv
    if os.path.isdir(source_fullpath + "\\Combined\\FUR") and os.path.isdir(source_fullpath + "\\Combined\\WU"):
        for root,dirs,files in os.walk(source_fullpath + "\\Combined\\FUR"):
            for name in files:
                if name.find(".bin") != -1 or name.find(".inf") != -1:
                    copy(source_fullpath + "Combined\\FUR\\" + name, target_fullpath + "\\HPFWUPDREC")
                    print(sourceFolder + "\\Combined\\FUR\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
        for root,dirs,files in os.walk(source_fullpath + "\\Combined\\WU"):
            for name in files:
                copy(source_fullpath + "\\Combined\\WU\\" + name, target_fullpath + "\\Capsule\\Windows")
                print(sourceFolder + "\\Combined\\WU\\" + name + " to " + targetFolder + "\\Capsule\\Windows" + " Copy succeeded.")
    # If Linux folder exist, copy files.
    if os.path.isdir(source_fullpath+"\\Combined\\Linux"):
        for root,dirs,files in os.walk(source_fullpath+"\\Combined\\Linux"):
            for name in files:
                copy(source_fullpath + "\\Combined\\Linux\\" + name, target_fullpath + "\\Capsule\\Linux")
                print(sourceFolder + "\\Combined\\Linux\\" + name + " to " + targetFolder + "\\Capsule\\Linux" + " Copy succeeded.")
    #=======For G4 other Fv
    elif os.path.isfile(source_fullpath + "\\Combined\\fwu.pfx"):
        for root,dirs,files in os.walk(source_fullpath + "\\Combined"):
            for name in files:
                if name.find(".bin") != -1 or name.find(".inf") != -1:
                    copy(source_fullpath + "\\Combined\\" + name, target_fullpath + "\\HPFWUPDREC")
                    print(sourceFolder + "\\Combined\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
                copy(source_fullpath + "\\Combined\\" + name, target_fullpath + "\\Capsule")
                print(sourceFolder + "\\Combined\\" + name + " to " + targetFolder + "\\Capsule" + " Copy succeeded.")
    # Bin file copy to FPTW&Global.
    for name in os.listdir(source_fullpath):# Bin file copy to FPTW&Global
        if (Platform_Flag(name) == "Q26") or (Platform_Flag(name) == "Q27") or \
            (Platform_Flag(name) == "R26") or (Platform_Flag(name) == "R24") or \
            (Platform_Flag(name) == "S25") or (Platform_Flag(name) == "S27") or (Platform_Flag(name) == "S29") or \
            (Platform_Flag(name) == "T25") or (Platform_Flag(name) == "T26") or (Platform_Flag(name) == "T27"): # 78 AMD 78 R26 78787878787878
            if name.find(NewVersion + "_16.bin") != -1:
                copy(source_fullpath + name, target_fullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" +" Copy succeeded.")
            if name.find(NewVersion + "_32.bin") != -1:
                copy(source_fullpath + name, target_fullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" +" Copy succeeded.")
        else:
            if name.find(NewVersion + ".bin") != -1:
                copy(source_fullpath + name, target_fullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" + " Copy succeeded.")
        if name.find(NewVersion + ".bin") != -1:
            copy(source_fullpath + name, target_fullpath + "\\Global\\BIOS")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.")
        if name.find(".xml") != -1:# Copy to XML
            if name.find(str(Platform_Flag(name))) != -1:
                copy(source_fullpath + name, target_fullpath + "\\XML")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\XML" + " Copy succeeded.")
        # For Smart flash copy *Pvt.bin.
        if name.find("Pvt.bin") != -1:
            copy(source_fullpath + name, target_fullpath)
            print(sourceFolder + "\\" + name + " to " + targetFolder + " Copy succeeded.")
    print()


# Find Fv Folder, Add to Match_list
def FindFvFolder(ProcessProjectList, NewVersion, NewBuildID):
    Match_list = []
    for Process in ProcessProjectList:
        for Dir in os.listdir(".\\"):
            Match = "Fv_" + Process + "_" + NewVersion + "_"
            if not NewBuildID == "":
                Match = "Fv_" + Process + "_" + NewVersion + "_" + NewBuildID
            if Match in Dir:
                if not str(Dir).find(".zip") != -1 and not str(Dir).find(".7z") != -1:
                    Match_list.append(Dir)
    return Match_list


# Find Fv Zip file
def FindFvZip(ProcessProjectList, ProjectNameInfo, NewVersion, NewBuildID):
    Match_list = []
    Github_PkgVersion = "FFFFFF"

    #======For New Github Fv
    if len(NewVersion) >= 6:
        Github_PkgVersion = f".{int(NewVersion[0:2])}.{int(NewVersion[2:4])}.{int(NewVersion[4:6])}"

    for i, project in enumerate(ProcessProjectList):
        # Construct matching strings based on old and new versions
        Match = f"Fv_{project}_{NewVersion}_{NewBuildID}" if NewBuildID else f"Fv_{project}_{NewVersion}_"

        for Dir in os.listdir(".\\"):
            # Check if it is an old version of Fv file
            if Match in Dir and (Dir.endswith(".zip") or Dir.endswith(".7z")):
                if Dir.endswith(".7z"):
                    new_name = f"{os.path.splitext(Dir)[0]}.zip"
                    os.rename(f".\\{Dir}", f".\\{new_name}")
                    Dir = new_name
                Match_list.append(Dir)

            # Checking for Github Fv Files
            if Github_PkgVersion in Dir and (project.lower() in Dir or project in Dir):
                # For U23 Github Fv project name mistake rename
                if not Dir.startswith(project.lower() + ProjectNameInfo[i].lower()):
                    new_name = project.lower() + ProjectNameInfo[i].lower() + Dir[len(project):]
                    os.rename(f".\\{Dir}", f".\\{new_name}")
                    Dir = new_name

                # Rename .nupkg and .7z to .zip
                if Dir.endswith(".nupkg") or Dir.endswith(".7z"):
                    new_name = f"{os.path.splitext(Dir)[0]}.zip"
                    os.rename(f".\\{Dir}", f".\\{new_name}")
                    Match_list.append(new_name)
                elif Dir.endswith(".zip"):
                    Match_list.append(Dir)

    return Match_list