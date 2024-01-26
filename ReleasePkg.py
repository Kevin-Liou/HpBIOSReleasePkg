#=================================
# coding=UTF-8
# Make Release Pkg Script.
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com

#This script is for making release pkg
#Now can use Intel G4, G5, G6, G8, G9, G10 AMD G4, G5, G6, G8 platform
#=================================
import os, logging, re
from colorama import Fore
from shutil import move, rmtree

from ReleasePkgLib import *

# In add the Projects to be here, and the following Board ID should also add.
BoardID=["P10","Q10~11","Q21~23","Q26~27","R11","R21~24","R26","S10~11","S21~23","S25~S29", "T11,T21~22", "T25~27", "U11,U21~23", "V11, V21~V23"]
ProjectName={"Q26":"ScottyRr", "Q27":"Scotty", "R24":"Worf", "R26":"Riker", "S25":"DoppioPco", "S27":"DoppioRn", "S29":"CubanoRn", "T25":"DoppioCzn", "T26":"CubanoCzn", "T27":"DoppioR8"}

Version_file_list=["BUFF2.nsh", "Buff2All.nsh", "Update32.bat", "Update64.bat", "UpdateEFI.nsh",
                "Update32_vPro.bat", "Update64_vPro.bat", "UpdateEFI_vPro.nsh"]

Not_Remove_file_rule=["Note", "note", "History", "How to Flash", "AMT_CFG", "logo", "sign.bin", "HPSignME", "Batch", "PD_FW"] # Priority over "Remove_file_rule"

Remove_file_rule=["DCI..+", ".cer", ".pfx", ".pvk", ".xlsm", ".log", r"Pvt.bin", r"metainfo.xml", r"Build.Log",
                r"\d\d_\d\d_\d\d.bin", r"\d\d_\d\d_\d\d.cat", r"\d\d_\d\d_\d\d.inf",
                r"\d{6}.bin", r"\d{6}.cat", r"\d{6}.cab", r"\d{6}.inf",
                r"\d{4}_12.bin", r"\d{4}_16.bin", r"\d{4}_32.bin",
                r"\w\d{2}_\d{4}.bin", r"\w\d{2}_\d{4}.cat", r"\w\d{2}_\d{4}.xml",r"\w\d{2}_\d{4}.inf", r"\w\d{2}_\d{6}.xml",
                r"P00\w{3}-\w{3}.zip", r"TBT_RTD3.+", r"HP_\w+_\w+_\w+_\w+_\w+_\w+_\d+.+", r"QA's report", r"QAreport",
                r"ME_+[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]+.bin"]

ProductionReleaseServer={   "type":"Production",
                            "host":"ftp.usa.hp.com",
                            "username":"sign_ron",
                            "pwd":"7gg9*0UA"}

TestReleaseServer={         "type":"Test",
                            "host":"ftp.usa.hp.com",
                            "username":"bios15ws",
                            "pwd":"e.QV9ra}"}

    # Script Start
if __name__ == '__main__':
    main_init()
    #=================Script Start==========================================================
    logging.debug("Debug Mode")
    print(("Make Release Pkg Script  " + Version()).center(90, "="))
    print("Input Information".center(90, "="))
    print("This script for making release pkg.\nNow can use Intel G4, G5, G6, G8, G9, G10&AMD G4, G5, G6, G8 platform.\n")
    OldVersion = InputStr("OldVersion:")# Input OldVersion & NewVersion & NewBuildID
    if OldVersion == "":
        ExitProgram("\nPlease Input OldVersion.")
    NewVersion = InputStr("NewVersion:")
    if NewVersion == "":
        ExitProgram("\nPlease Input NewVersion.")
    NewBuildID = InputStr("NewBuildID:")# ex: 020106_"0001"
    ProcessProject = InputStr("                                   (Can Multiple Select)\
        \nPlease Enter Projects To Processed"+str(BoardID)+":")# Input need Process boardID, Can multiple choice
    ProcessProjectList = ProcessProject.upper().split() # ex:['U21', 'U23']
    CheckBiosVersion(OldVersion, NewVersion, NewBuildID, ProcessProject) # Check Bios Version
    if ProcessProject == "":
        ExitProgram("\nPlease Input Project.")

    #=================Find Need Process Old Pkg==============================================
    print("Find Need Process Old Pkg".center(90, "="))
    OldBuildID = "0"; NeedProcOldPkg = []
    print("Start Find Process Project Old Pkg.")
    for Project in ProcessProjectList:
        temp = []
        #======For Intel
        if (Platform_Flag(Project) == "Intel G3") or (Platform_Flag(Project) == "Intel G4") or \
            (Platform_Flag(Project) == "Intel G5") or (Platform_Flag(Project) == "Intel G6") or \
            (Platform_Flag(Project) == "Intel G8") or (Platform_Flag(Project) == "Intel G9") or \
            (Platform_Flag(Project) == "Intel G10"):
            for Dir in os.listdir(".\\"):# Find old version pkg.
                if not Dir.split("_")[0] == "Fv" and not Dir.find(".7z") != -1 and not Dir.find(".zip") !=- 1:
                    if Project + "_" + OldVersion in Dir:
                        temp.append(Dir)
            if len(temp) > 1 and OldBuildID == "0":# If find old version pkg have buildID.
                OldBuildID = InputStr("\n" + str(temp) + "\n" + Project + "_" + OldVersion + " Please enter the BuildID of the package you want to select:")
                if OldBuildID == "" or OldBuildID == "0000":
                    temp = [a for a in temp if len(a.split("_")) == 4]
                else:
                    temp = [a for a in temp if Project + "_" + OldVersion + "_" + OldBuildID in a]
                if not len(temp) == 1:
                    temp = [Dir for Dir in os.listdir(".\\") if Project + "_" + OldVersion in Dir]
                    OldBuildID = InputStr("\n" + str(temp) + "\n" + Project + "_" + OldVersion + " Please enter the BuildID of the package you want to select:")
                    if OldBuildID == "" or OldBuildID == "0000":
                        temp = [a for a in temp if len(a.split("_")) == 4]
                    else:
                        temp = [a for a in temp if Project + "_" + OldVersion + "_" + OldBuildID in a]
            if len(temp) == 0:# Can't find old version pkg.
                ExitProgram(Project + "_" + OldVersion + " Old Pkg folder can't find, Please check.")
            else:# Add find old version pkg.
                NeedProcOldPkg.append(temp[0])
        #======For AMD
        else:
            OldVersion_AMD = OldVersion[0:2] + "." + OldVersion[2:4] + "." + OldVersion[4:6] # ex.['Q26', '02.07.03']
            if (Platform_Flag(Project) == "R26") or (Platform_Flag(Project) == "S25") or \
                (Platform_Flag(Project) == "S27") or (Platform_Flag(Project) == "S29") or \
                (Platform_Flag(Project) == "T25") or (Platform_Flag(Project) == "T26") or \
                (Platform_Flag(Project) == "T27") :# 78 AMD 78 R26 S25~S29 T25~T27
                OldVersion_AMD = OldVersion
            for Dir in os.listdir(".\\"):# Find old version pkg.
                if not Dir.split("_")[0] == "Fv" and not Dir.find(".7z") != -1 and not Dir.find(".zip") != -1:
                    if Project + "_" + OldVersion_AMD in Dir:
                        temp.append(Dir)
            if len(temp) > 1 and OldBuildID == "0":# If find old version pkg have buildID.
                OldBuildID = InputStr("\n" + str(temp) + "\n" + Project + "_" + OldVersion_AMD + " Please enter the BuildID of the package you want to select:")
                if OldBuildID == "" or OldBuildID == "0000":
                    temp = [a for a in temp if len(a.split("_")) == 2]
                else:
                    temp = [a for a in temp if Project + "_" + OldVersion_AMD + "_" + OldBuildID in a]
                if not len(temp) == 1:
                    temp = [Dir for Dir in os.listdir(".\\") if Project + "_" + OldVersion_AMD in Dir]
                    OldBuildID = InputStr("\n" + str(temp) + "\n" + Project + "_" + OldVersion_AMD + " Please enter the BuildID of the package you want to select:")
                    if OldBuildID == "" or OldBuildID == "0000":
                        temp = [a for a in temp if len(a.split("_")) == 2]
                    else:
                        temp = [a for a in temp if Project + "_" + OldVersion_AMD + "_" + OldBuildID in a]
            if len(temp) == 0:# Can't find old version pkg.
                ExitProgram(Project + "_" + OldVersion_AMD + " Old Pkg folder can't find, Please check.")
            else:# Add find old version pkg.
                NeedProcOldPkg.append(temp[0])
    print("\nYour need process old Pkg:\n" + str(NeedProcOldPkg))

    #=================Make need process old/new Pkg info table===============================
    OldProcPkgInfo = []
    OldProcPkgInfo = [Proc.split("_") for Proc in NeedProcOldPkg]# Split to process
    NewProcPkgInfo = [Proc.split("_") for Proc in NeedProcOldPkg]# New Pkg Name List
    ProjectNameInfo = [Proc.split("_")[0] for Proc in NeedProcOldPkg]# ex:['Pacman', 'Asteroid']
    for OldProcPkg in OldProcPkgInfo:
        #======For Intel
        if (Platform_Flag(OldProcPkg) == "Intel G3") or (Platform_Flag(OldProcPkg) == "Intel G4") or \
            (Platform_Flag(OldProcPkg) == "Intel G5") or (Platform_Flag(OldProcPkg) == "Intel G6") or \
            (Platform_Flag(OldProcPkg) == "Intel G8") or (Platform_Flag(OldProcPkg) == "Intel G9") or \
            (Platform_Flag(OldProcPkg) == "Intel G10"):
            for Proc in range(len(NewProcPkgInfo)):
                NewProcPkgInfo[Proc] = NewProcPkgInfo[Proc][:3] # ex:['Harp', 'MV', 'Q21']
                NewProcPkgInfo[Proc].append(NewVersion) # ex:['Harp', 'MV', 'Q21', 'NewVersion']
                NewBuildID = "" if NewBuildID == "0000" else NewBuildID
                if NewBuildID: NewProcPkgInfo[Proc].append(NewBuildID)
            break #Otherwise it will run more times
        #======For AMD
        else:
            NewVersion_AMD = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] # ex.['Q26', '02.07.03']
            for Proc in range(len(NeedProcOldPkg)):
                if (Platform_Flag(OldProcPkg) == "R26") or (Platform_Flag(OldProcPkg) == "S25") or \
                    (Platform_Flag(OldProcPkg) == "S27") or (Platform_Flag(OldProcPkg) == "S29") or \
                    (Platform_Flag(OldProcPkg) == "T25") or (Platform_Flag(OldProcPkg) == "T26") or \
                    (Platform_Flag(OldProcPkg) == "T27") :# 78 AMD 78 R26 S25~S29 T25~T27
                    NewVersion_AMD = NewVersion # ex.'020703'
                    if len(NewProcPkgInfo[Proc]) == 3:
                        del NewProcPkgInfo[Proc][2]
                    NewProcPkgInfo[Proc][1] = NewVersion_AMD # ex:['Q26', 'NewVersion_AMD(02.07.03)']
                if (Platform_Flag(OldProcPkg) == "R24"):
                    NewProcPkgInfo[Proc] = NewProcPkgInfo[Proc][:2] # ex:['Worf', 'R24']
                    NewProcPkgInfo[Proc].append(NewVersion)
                    NewProcPkgInfo[Proc][2] = NewVersion_AMD # ex:['Worf', 'R24','NewVersion_AMD(02.07.03)']
            NewBuildID = "" if NewBuildID == "0000" else NewBuildID
            if NewBuildID: NewProcPkgInfo[Proc].append(NewBuildID)
            break #Otherwise it will run more times
    logging.debug("OldProcPkgInfo:" + str(OldProcPkgInfo))
    logging.debug("NewProcPkgInfo:" + str(NewProcPkgInfo))

    #=================Find Fv Folder Or Zip File=============================================
    print("Find Fv Folder Or Zip File".center(90, "="))
    Match_folder_list = FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)
    Match_zip_list = FindFvZip(ProcessProjectList, ProjectNameInfo, NewVersion, NewBuildID)

    for PkgInfo in range(len(NewProcPkgInfo)):
        #======For AMD Start
        if (Platform_Flag(OldProcPkgInfo) == "Q26") or (Platform_Flag(OldProcPkgInfo) == "Q27") or (Platform_Flag(OldProcPkgInfo) == "R26") or \
            (Platform_Flag(OldProcPkgInfo) == "S25") or (Platform_Flag(OldProcPkgInfo) == "S27") or \
            (Platform_Flag(OldProcPkgInfo) == "S29") or (Platform_Flag(OldProcPkgInfo) == "T25") or \
            (Platform_Flag(OldProcPkgInfo) == "T26") or (Platform_Flag(OldProcPkgInfo) == "T27") :# ex.['Q26', '01.04.01']=>['ScottyRr', 'ScottyRr', 'Q26', '010401']
            NewProcPkgInfo[PkgInfo][1] = NewVersion
            NewProcPkgInfo[PkgInfo].insert(0, ProjectName[NewProcPkgInfo[PkgInfo][0]])
            NewProcPkgInfo[PkgInfo].insert(0, ProjectName[NewProcPkgInfo[PkgInfo][1]])
        if (Platform_Flag(OldProcPkgInfo) == "R24"): # ex.['Worf', 'R24', '02.07.03']=>['Worf', 'Worf', 'R24', '020703']
            NewProcPkgInfo[PkgInfo][2] = NewVersion
            NewProcPkgInfo[PkgInfo].insert(0, NewProcPkgInfo[PkgInfo][0]) # add "Worf"
        #======For AMD End

    print("Your Fv Folder: %s" % str(Match_folder_list))
    print("Your Fv Zip File: %s" % str(Match_zip_list))
    # If can't find Fv folder or Zip file.
    if len(Match_folder_list) == 0 and len(Match_zip_list) == 0:
        print("Can't find Fv folder and zip file.\nDownload Fv files from Production Release FTP.\n")
        temp = Ftp_multi(NewProcPkgInfo, ProductionReleaseServer, TestReleaseServer)[:] # Download Fv files from Production Release FTP
        for name in temp:
            if str(name).find(".zip") != -1:
                Match_zip_list.append(name)
    # Number of Fv folders not match Project list
    elif len(Match_folder_list) < len(ProcessProjectList) and len(Match_folder_list) != 0 and len(Match_zip_list) < len(ProcessProjectList):
        print("Number of Fv folders not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp = Ftp_multi(NewProcPkgInfo, ProductionReleaseServer, TestReleaseServer)[:] # Download Fv files from Production Release FTP
        for name in temp:
            if str(name).find(".zip") != -1:
                Match_zip_list.append(name)
    # Can't find Fv folders and Number of Fv Zip files not match Project list
    elif len(Match_folder_list) == 0 and len(Match_zip_list) < len(ProcessProjectList):
        print("Fv Zip files not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp = Ftp_multi(NewProcPkgInfo, ProductionReleaseServer, TestReleaseServer)[:] # Download Fv files from Production Release FTP
        for name in temp:
            if str(name).find(".zip") != -1:
                Match_zip_list.append(name)
    # If Fv folder already exists
    elif len(Match_folder_list) == len(ProcessProjectList):
        print("Fv folder already exists.\n")

    Match_folder_list = FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)
    if len(Match_folder_list) > 1:
        ExitProgram("Find more than one fv folder, Please check.")

    # Find Fv Zip file and can't find Fv folder start extracting
    if Match_zip_list and not Match_folder_list:
        if len(Match_zip_list) > 1:
            ExitProgram("Find more than one fv zip file, Please check.")
        print("\nFind Fv Zip File, Start Extracting.")
        for i, zip_file in enumerate(Match_zip_list):
            Foldername = zip_file.replace(".zip", "")

            # Determine if the file needs to be processed
            is_github_package = ProjectNameInfo[i].lower() in zip_file.lower() and "fv_" not in zip_file.lower()
            is_normal_package = "fv_" in zip_file.lower()
            logging.debug(f'is_github_package: {is_github_package}, is_normal_package: {is_normal_package}')

            # Skip if directory already exists
            if os.path.isdir(".\\" + Foldername):
                print(f"Fv folder {Foldername} already exists.")
                #if os.path.isfile(".\\" + zip_file):
                #    os.remove(".\\" + zip_file)
                continue

            try:
                UnZip(zip_file)

                # For GitHub package file
                if is_github_package:
                    new_folder_name = f"Fv_{ProcessProjectList[i]}_{NewVersion}_{NewBuildID if NewBuildID else '0000'}"
                    SafeRename(f".\\{Foldername}\\Fv", f".\\{Foldername}\\{new_folder_name}")
                    if not os.path.isdir(f".\\{new_folder_name}") and os.path.isdir(f".\\{Foldername}\\{new_folder_name}"):
                        move(f".\\{Foldername}\\{new_folder_name}", ".\\")
                        if os.path.isdir(f".\\{new_folder_name}"):
                            rmtree(f".\\{Foldername}")
                            Match_folder_list.append(new_folder_name)
                    print(f"Github package {zip_file} Extract succeeded.")

                # For normal package file
                elif is_normal_package:
                    Match_folder_list.append(Foldername)
                    print(f"Normal package {zip_file} Extract succeeded.")

            except Exception as e:
                print(f"Error processing {zip_file}: {e}")

        print(f"\nNow Your Fv Folder: {Match_folder_list}")
    else:
        print(f"\nNow Your Fv Folder: {Match_folder_list}")

    # Working with multiple folders
    MatchMultipleFolder(Match_folder_list, ProcessProjectList, NewVersion)

    #=================Find New Pkg Or Add New Pkg=============================================
    print("Find New Pkg Or Add New Pkg".center(90, "="))

    for PkgInfo in range(len(NewProcPkgInfo)):
        #======For AMD Start
        if (Platform_Flag(OldProcPkgInfo) == "Q26") or (Platform_Flag(OldProcPkgInfo) == "Q27") or (Platform_Flag(OldProcPkgInfo) == "R26") or \
            (Platform_Flag(OldProcPkgInfo) == "S25") or (Platform_Flag(OldProcPkgInfo) == "S27") or \
            (Platform_Flag(OldProcPkgInfo) == "S29") or (Platform_Flag(OldProcPkgInfo) == "T25") or \
            (Platform_Flag(OldProcPkgInfo) == "T26") or (Platform_Flag(OldProcPkgInfo) == "T27") : # ['ScottyRr', 'ScottyRr', 'Q26', '010401']=>['Q26', '01.04.01']
            NewProcPkgInfo[PkgInfo].remove(ProjectName[NewProcPkgInfo[PkgInfo][2]])
            NewProcPkgInfo[PkgInfo].remove(ProjectName[NewProcPkgInfo[PkgInfo][1]])
            if (Platform_Flag(OldProcPkgInfo) == "Q26") or (Platform_Flag(OldProcPkgInfo) == "Q27"):
                NewProcPkgInfo[PkgInfo][1] = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
        if (Platform_Flag(OldProcPkgInfo) == "R24"): # ['Worf', 'Worf', 'R24, '010401']=>['Worf', 'R24', '01.04.01']
            NewProcPkgInfo[PkgInfo].remove(NewProcPkgInfo[PkgInfo][0]) # remove 'Worf'
            NewProcPkgInfo[PkgInfo][1] = Platform_Flag(OldProcPkgInfo)
            NewProcPkgInfo[PkgInfo][2] = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
        #======For AMD End

    for OProc in range(len(OldProcPkgInfo)):# How much Old Version folder
        OldVersionPath = ".\\" + ("_").join(OldProcPkgInfo[OProc])
        NewVersionPath = ".\\" + ("_").join(NewProcPkgInfo[OProc])
        if not os.path.isdir(NewVersionPath):# Check NewVersion Folder Exist
            if not os.path.isdir(OldVersionPath + "\\" + ("_").join(OldProcPkgInfo[OProc])):# Check Old Pkg is in folder???
                Copy_Release_Folder(OldVersionPath, NewVersionPath)
            elif not os.path.isdir(OldVersionPath + "\\FPTW"):
                if os.path.isdir(OldVersionPath + "\\" + ("_").join(OldProcPkgInfo[OProc])):# Check Old Pkg is in folder???
                    Copy_Release_Folder(OldVersionPath + "\\" + ("_").join(OldProcPkgInfo[OProc]), NewVersionPath)
                else:
                    print("Pkg " + ("_").join(OldProcPkgInfo[OProc]) + " can't find.")
            else:
                print("Pkg " + ("_").join(OldProcPkgInfo[OProc]) + " can't find.")
        else:
            ExitProgram("Pkg " + NewVersionPath.split("\\")[-1] + Fore.RED + " already exists.")
    if len(Match_folder_list) == 0:
        ExitProgram("Can't find anything Fv folder.\n")

    #=================Modify Pkg Update Version==============================================
    print("Modify Pkg Update Version".center(90, "="))
    BiosBuildDate = CheckBiosBuildDate(Match_folder_list)
    BiosMrcVersion = GetMrcVersion(Match_folder_list)
    BiosIshVersion = GetIshVersion(Match_folder_list)
    BiosPmcVersion = GetPmcVersion(Match_folder_list)
    BiosNphyVersion = GetNphyVersion(Match_folder_list)
    BiosBinaryChecksum = CheckFileChecksum(Match_folder_list, NewVersion)
    logging.debug(f'NewProcPkgInfo: {NewProcPkgInfo}')
    for NProc in NewProcPkgInfo:# Pkg Modify Update Version
        Path = os.getcwd() + "\\" + ("_").join(NProc)
        ReleaseNoteName = ""
        ReleaseNote_docx = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("Release" in ReleaseNote) and (".docx" in ReleaseNote)]
        ReleaseNote_xlsm = [ReleaseNote for ReleaseNote in os.listdir(Path)
                            if  ("Release" in ReleaseNote)
                            and ("Note" in ReleaseNote)
                            and ReleaseNote.endswith(".xlsm")
                            and not ReleaseNote.startswith("~$")] # Ignore temp file
        logging.debug(f'ReleaseNote_docx: {ReleaseNote_docx}')
        logging.debug(f'ReleaseNote_xlsm: {ReleaseNote_xlsm}')
        #======For Intel
        if (Platform_Flag(NProc) == "Intel G3") or (Platform_Flag(NProc) == "Intel G4") or \
            (Platform_Flag(NProc) == "Intel G5") or (Platform_Flag(NProc) == "Intel G6") or \
            (Platform_Flag(NProc) == "Intel G8") or (Platform_Flag(NProc) == "Intel G9"):
            if os.path.isdir(Path + "\\FPTW"):# Check Folder Exist
                # If get release note G4
                if len(ReleaseNote_docx) == 1:
                    os.chdir(Path)
                    os.rename(ReleaseNote_docx[0], ("_").join(NProc) + " release note.docx")
                    if os.path.isfile(("_").join(NProc) + " release note.docx"):
                        print("ReleaseNote Rename to " + ("_").join(NProc) + " release note.docx" + " succeeded.")
                # If get release note G5 and late (Support 2.0 release note)
                elif len(ReleaseNote_xlsm) > 0:
                    if len(ReleaseNote_xlsm) > 1:
                        print("\n" + "Please choose which release note you want to modify\n")
                        for index, release_note in enumerate(ReleaseNote_xlsm, start=1):
                            print(f"{index}. {release_note}")
                        ReleaseNoteName = ReleaseNote_xlsm[(int(input("\nPlease enter the number:"))) - 1]
                    else:
                        ReleaseNoteName = ReleaseNote_xlsm[0]

                    # If get release note G5 and late
                    if ReleaseNoteName:
                        os.chdir(Path)
                        # Rename release note
                        VersionPattern = re.compile(r'\d{6}(?:_\d{4})?')
                        ReleaseNewVersion = NewVersion if NewBuildID == "" else NewVersion + "_" + NewBuildID
                        if ReleaseNewVersion in ReleaseNoteName:
                            print("ReleaseNote Already Rename.")
                            break
                        else:
                            NewReleaseNoteName = VersionPattern.sub(ReleaseNewVersion, ReleaseNoteName, 1)
                            os.rename(ReleaseNoteName, NewReleaseNoteName)
                            print(f"Renamed {ReleaseNoteName} to {NewReleaseNoteName}")
                            ReleaseNoteName = NewReleaseNoteName
                        # Modify release note
                        if os.path.isfile(ReleaseNoteName):
                            ModifyReleaseNote(NProc, ReleaseNoteName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
                            print(f"ReleaseNote {ReleaseNoteName} Modify succeeded.")
                else:
                    print("Can't find release note file.")
                    ExitProgram(f"ReleaseNote_xlsm: {ReleaseNote_xlsm}, ReleaseNote_docx: {ReleaseNote_docx}")

                os.chdir(Path + "\\FPTW")
                ChangeBuildID(NProc, Version_file_list, NewVersion) # Modify BIOS version in Version_file_list files.
            else:
                ExitProgram("Pkg Folder " + ("_").join(NProc) + " can't find.\n")

        #======For AMD (Not use need to modify)
        else:
            if os.path.isdir(Path + "\\AMDFLASH"):# Check Folder Exist
                if len(ReleaseNote_docx) == 1: # If get release note G4
                    os.chdir(Path)
                    os.rename(ReleaseNote_docx[0], "Scotty_" + ("_").join(NProc) + "_Release_Notes.docx")
                    if os.path.isfile("Scotty_" + ("_").join(NProc) + "_Release_Notes.docx"):
                        print("ReleaseNote Rename to " + "Scotty_" + ("_").join(NProc) + "_Release_Notes.docx" + " succeeded.")
                elif len(ReleaseNote_xlsm) == 1: # If get release note G5 and late
                    os.chdir(Path)
                    ReleaseName = ReleaseNote_xlsm[0].split("Note_")[0]
                    if (NewBuildID == "" or NewBuildID == "0000"):
                        if ReleaseNote_xlsm[0] == ReleaseName + "Note_" + NewVersion + ".xlsm":
                            print("ReleaseNote Already Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + "Note_" + NewVersion + ".xlsm")
                        if os.path.isfile(ReleaseName + "Note_" + NewVersion + ".xlsm"):
                            ReleaseNoteName = ReleaseName + "Note_" + NewVersion + ".xlsm"
                            ModifyReleaseNote(NProc, ReleaseNoteName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
                            print("ReleaseNote Rename to " + ReleaseName + "Note_" + NewVersion + ".xlsm" + " succeeded.")
                    else:
                        if ReleaseNote_xlsm[0] == ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm":
                            print("ReleaseNote Already Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm")
                        if os.path.isfile(ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm"):
                            ReleaseNoteName = ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm"
                            ModifyReleaseNote(NProc, ReleaseNoteName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
                            print("ReleaseNote Rename to " + ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm" + " succeeded.")
                else:
                    print("Can't find release note file or more than one release note.")
                    ExitProgram(f"ReleaseNote_xlsm: {ReleaseNote_xlsm}")
                os.chdir(Path + "\\AMDFLASH")
                ChangeBuildID(NProc, Version_file_list, NewVersion)
            else:
                ExitProgram("Pkg " + ("_").join(NProc) + " can't find.\n")

        os.chdir("..\..")

    #=================Remove Pkg Old File=====================================================
    print("Remove Pkg Old File".center(90, "="))
    for NProc in NewProcPkgInfo:
        target_folder = ("_").join(NProc)
        if os.path.isdir(".\\"+target_folder):
            RemoveOldFileInDir(target_folder, Remove_file_rule, Not_Remove_file_rule)
        else:
            print("Pkg "+("_").join(NProc)+" can't find.\n")

    #=================Fv File Rename And Copy To Pkg===========================================
    print("Fv File Rename And Copy To Pkg".center(90, "="))
    logging.debug(f'Match_folder_list: {Match_folder_list}')
    logging.debug(f'NewProcPkgInfo: {NewProcPkgInfo}')
    for Fv in Match_folder_list:
        for NProc in NewProcPkgInfo:
            #======For Intel
            if (Platform_Flag(NProc) == "Intel G3") or (Platform_Flag(NProc) == "Intel G4") or \
                (Platform_Flag(NProc) == "Intel G5") or (Platform_Flag(NProc) == "Intel G6") or \
                (Platform_Flag(NProc) == "Intel G8") or (Platform_Flag(NProc) == "Intel G9") or \
                (Platform_Flag(NProc) == "Intel G10"):
                if Fv.split("_")[1] == NProc[2]:
                    if os.path.isdir(".\\"+Fv):
                        Path = ".\\" + Fv
                        Board_version = NProc[2]+"_" + NProc[3]
                        if os.path.isfile(Path + "\\" + Board_version + "_12.bin") or os.path.isfile(Path + "\\" + Board_version + "_16.bin"):# If Alreadly renamed
                            if os.path.isfile(Path+"\\"+Board_version+".xml"):
                                print(Board_version + "_12.bin or _16.bin & " + Board_version + ".xml alreadly renamed.")
                        if os.path.isfile(Path + "\\" + Board_version + ".bin") and os.path.isfile(Path + "\\" + Board_version + "_16.bin"):
                            os.rename(Path + "\\" + Board_version + ".bin", Path + "\\" + Board_version + "_12.bin")  # Rename Fv folder 2 files
                            os.rename(Path + "\\" + NProc[2] + ".xml", Path + "\\" + Board_version + ".xml")
                            print(Board_version + "_12.bin & " + Board_version + ".xml rename succeeded.")
                        if os.path.isfile(Path + "\\" + Board_version + ".bin") and os.path.isfile(Path + "\\" + Board_version + "_32.bin"):
                            os.rename(Path + "\\" + Board_version + ".bin", Path + "\\" + Board_version + "_16.bin")  # Rename Fv folder 2 files
                            os.rename(Path + "\\" + NProc[2] + ".xml", Path + "\\" + Board_version + ".xml")
                            print(Board_version + "_16.bin & " + Board_version + ".xml rename succeeded.")
                        if (os.path.isfile(Path+"\\"+Board_version+"_32.bin") or os.path.isfile(Path+"\\"+Board_version+"_16.bin")) \
                            and os.path.isdir(".\\"+("_").join(NProc)):# Check Pkg Folder Exist
                            Copy_Release_Files(Fv, ("_").join(NProc), NProc, Match_folder_list)
                        else:
                            print("Pkg " + ("_").join(NProc) + " can't find.")
                    else:
                        print("Need to be processed Fv folder:" + Fv + " can't find.\n")
            #======For AMD
            else:
                if (Platform_Flag(NProc) == "R26") or (Platform_Flag(NProc) == "R24") or (Platform_Flag(NProc) == "Q26") or (Platform_Flag(NProc) == "Q27") or \
                    (Platform_Flag(NProc) == "S25") or (Platform_Flag(NProc) == "S27") or (Platform_Flag(NProc) == "S29") or \
                    (Platform_Flag(NProc) == "T25") or (Platform_Flag(NProc) == "T26") or (Platform_Flag(NProc) == "T27") :
                    if (Fv.split("_")[1] == NProc[0]) or (Fv.split("_")[1] == NProc[1]):
                        if os.path.isdir(".\\" + Fv):
                            Path = os.getcwd() + "\\" + Fv
                            Board_version = NProc[0] + "_" + NewVersion
                        if (Platform_Flag(NProc) == "R24"):
                            Board_version = NProc[1] + "_" + NewVersion
                        if os.path.isfile(Path + "\\" + Board_version + ".bin") or os.path.isfile(Path + "\\" + Board_version + "_16.bin") or os.path.isfile(Path + "\\" + Board_version + "_32.bin"):# For 16MB BIOS
                            if os.path.isfile(Path + "\\" + Board_version[:3] + ".xml"):
                                if os.path.isdir(".\\" + ("_").join(NProc)):# Check Pkg Folder Exist
                                    Copy_Release_Files_AMD(Fv, ("_").join(NProc), NewVersion)
                                else:
                                    print("Pkg " + ("_").join(NProc) + " can't find.")
                        print(Board_version + "_16.bin & " + Board_version + ".xml rename succeeded.")
    if len(Match_folder_list) == 0 or len(ProcessProjectList) == 0:
        print("Can't find anything Fv folder.\n")

    #=================Check Tool Version Is Match In Table======================================
    print("Check Tool Version Is Match In Table".center(90, "=")+"\n")
    try:
        for NProc in NewProcPkgInfo:
            #======For Intel G5&G6 and late
            if ((Platform_Flag(NProc) == "Intel G5") or (Platform_Flag(NProc) == "Intel G6") or \
                (Platform_Flag(NProc) == "Intel G8") or (Platform_Flag(NProc) == "Intel G9") or \
                (Platform_Flag(NProc) == "Intel G10")) and \
                os.path.isdir(".\\" + ("_").join(NProc) + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)"):
                Tool_version_table_path = ".\\" + ("_").join(NProc) + "\\FactoryUtility\\ToolVersion.xlsx"
                Tool_version_info = ReadToolVersionTable(Tool_version_table_path)
                Check = "Match"
                for name, type, verinfo, path, note in Tool_version_info:
                    ver = ChangeVersionInfo(verinfo)
                    #date = ChangeDataInfo(dateinfo)
                    CompareInfo(NProc, name, ver, path, Tool_version_table_path)
            #======For ADM G5 and late
            elif ((Platform_Flag(NProc) == "R26") or (Platform_Flag(NProc) == "S25") or \
                (Platform_Flag(NProc) == "S27") or (Platform_Flag(NProc) == "S29") or \
                (Platform_Flag(NProc) == "T25") or (Platform_Flag(NProc) == "T26") or\
                (Platform_Flag(NProc) == "T27") ) and \
                os.path.isdir(".\\" + ("_").join(NProc) + "\\Capsule\\Windows"):
                Tool_version_table_path = ".\\" + ("_").join(NProc) + "\\FactoryUtility\\ToolVersion.xlsx"
                Tool_version_info = ReadToolVersionTable(Tool_version_table_path)
                Check = "Match"
                for name, type, verinfo, path, note in Tool_version_info:
                    ver = ChangeVersionInfo(verinfo)
                    CompareInfo(NProc, name, ver, path, Tool_version_table_path)
    except ValueError:
        pass
        #print("\n"+"Table Format "+Fore.RED+"Error"+".\n")
    except:
        pass
        #print("\n"+"Check Toolversion "+Fore.RED+"Error"+".\n")

    #=================End=======================================================================
    print("End".center(90, "="))
    #======For Intel Check
    if (Platform_Flag(OldProcPkgInfo) == "Intel G3") or (Platform_Flag(OldProcPkgInfo) == "Intel G4") or \
        (Platform_Flag(OldProcPkgInfo) == "Intel G5") or (Platform_Flag(OldProcPkgInfo) == "Intel G6") or \
        (Platform_Flag(OldProcPkgInfo) == "Intel G8") or (Platform_Flag(OldProcPkgInfo) == "Intel G9") or \
        (Platform_Flag(OldProcPkgInfo) == "Intel G10"):
        CheckPkg(NewProcPkgInfo)# Check new release Pkg is OK?
        PrintBiosBuildDate(Match_folder_list, BiosBuildDate)
        PrintBiosBinaryChecksum(NewProcPkgInfo, BiosBinaryChecksum, NewVersion)
    #======For AMD Check
    else:
        CheckPkg_AMD(NewProcPkgInfo, NewVersion, NewBuildID)# Check new release Pkg is OK?
        PrintBiosBuildDate(Match_folder_list, BiosBuildDate)
        PrintBiosBinaryChecksum(NewProcPkgInfo, BiosBinaryChecksum, NewVersion)
    ExitProgram("\nFinally pkg please compare with leading project.\n")
    # =================Script End===============================================================