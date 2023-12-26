#=================================
# coding=UTF-8
# Make Release Pkg Script.
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com

# This script is for making release pkg tool
# Now can use Intel G4, G5, G6, G8, G9, G10 AMD G4, G5, G6, G8 platform
# This file is main code, other code in ReleasePkgLib folder
#=================================
import os
import sys
import logging
from colorama import Fore
from shutil import move, rmtree
from ReleasePkgLib import *


# Script Start
if __name__ == '__main__':
    args, Config_data = Main_init()
    #=================Input BIOS Information===================================================
    print(("Make Release Pkg Script  " + Version()).center(90, "="))
    print("This tool allows you to make new release packages.\n\nCurrently supported platforms:\n  Intel:G4, G5, G6, G8, G9, G10\n  AMD:G4, G5, G6, G8\n")
    if args == "Debug mode":
        OldVersion, NewVersion, NewBuildID, ProcessProject, ProcessProjectList = Config_debug()
    else:
        OldVersion = InputStr("Please input package information.\nOldVersion:") # Input OldVersion & NewVersion & NewBuildID
        if OldVersion == "":
            input("\nPlease Input OldVersion.(ex. 020200)")
            sys.exit()
        NewVersion = InputStr("NewVersion:")
        if NewVersion == "":
            input("\nPlease Input NewVersion.(ex. 020300)")
            sys.exit()
        NewBuildID = InputStr("NewBuildID:") # ex: 020106_"0001"
        ProcessProject = InputStr("                                   (You can choose more, ex. U21 U22 U23)\
            \nPlease Enter Projects To Processed"+str(Config_data['BoardID'])+":") # Input need Process boardID, Can multiple choice
        ProcessProjectList = ProcessProject.upper().split() # ex:['U21', 'U23']
        if ProcessProject == "":
            input("\nPlease Input Project.(ex. U21)")
            sys.exit()

    #=================Finding Old Packages To Process==============================================
    print("Finding Old Packages To Process".center(90, "="))
    OldBuildID = "0"
    NeedProcOldPkg = []
    print("Start looking for old packages.")
    for Project in ProcessProjectList:
        temp = []
        #======For Intel Project
        if Platform_Flag(Project) in Intel_Platforms:
            # Find old version pkg folder.
            for Dir in os.listdir(".\\"):
                if not Dir.split("_")[0] == "Fv" and not Dir.find(".7z") != -1 and not Dir.find(".zip") !=- 1:
                    if Project + "_" + OldVersion in Dir:
                        temp.append(Dir)
            # If the find old version pkg has a buildID.
            if len(temp) > 1 and OldBuildID == "0":
                OldBuildID = InputStr("\n"+str(temp) + "\n" + Project + "_" + OldVersion + " Please Select Old BuildID:")
                if OldBuildID == "" or OldBuildID == "0000":
                    temp = [a for a in temp if len(a.split("_")) == 4]
                else:
                    temp = [a for a in temp if Project + "_" + OldVersion + "_" + OldBuildID in a]
                if not len(temp) == 1:
                    temp = [Dir for Dir in os.listdir(".\\") if Project + "_" + OldVersion in Dir]
                    OldBuildID = InputStr("\n" + str(temp) + "\n" + Project + "_" + OldVersion + " Please Select OldBuildID:")
                    if OldBuildID == "" or OldBuildID == "0000":
                        temp = [a for a in temp if len(a.split("_")) == 4]
                    else:
                        temp = [a for a in temp if Project + "_" + OldVersion + "_" + OldBuildID in a]
            # If can't find old version pkg.
            if len(temp) == 0:
                input(Project + "_" + OldVersion + " Old Pkg folder can't find, Please check.")
                sys.exit()
            # Add find old version pkg.
            else:
                NeedProcOldPkg.append(temp[0])
        #======For AMD Project
        else:
            OldVersion_AMD = OldVersion[0:2] + "." + OldVersion[2:4] + "." + OldVersion[4:6] # ex.['Q26', '02.07.03']
            # For AMD special project setting.
            if Platform_Flag(Project) in AMD_Platforms_Special:
                OldVersion_AMD = OldVersion
            # Find old version pkg folder.
            for Dir in os.listdir(".\\"):
                if not Dir.split("_")[0] == "Fv" and not Dir.find(".7z") != -1 and not Dir.find(".zip") != -1:
                    if Project + "_" + OldVersion_AMD in Dir:
                        temp.append(Dir)
            # If the find old version pkg has a buildID.
            if len(temp) > 1 and OldBuildID == "0":
                OldBuildID = InputStr("\n" + str(temp) + "\n" + Project + "_" + OldVersion_AMD + " Please Select Old BuildID:")
                if OldBuildID == "" or OldBuildID == "0000":
                    temp = [a for a in temp if len(a.split("_")) == 2]
                else:
                    temp = [a for a in temp if Project + "_" + OldVersion_AMD + "_" + OldBuildID in a]
                if not len(temp) == 1:
                    temp = [Dir for Dir in os.listdir(".\\") if Project + "_" + OldVersion_AMD in Dir]
                    OldBuildID = InputStr("\n" + str(temp) + "\n" + Project + "_" + OldVersion_AMD + " Please Select OldBuildID:")
                    if OldBuildID == "" or OldBuildID == "0000":
                        temp = [a for a in temp if len(a.split("_")) == 2]
                    else:
                        temp = [a for a in temp if Project + "_" + OldVersion_AMD + "_" + OldBuildID in a]
            # If can't find old version pkg.
            if len(temp) == 0:
                print(Project + "_" + OldVersion_AMD + " Old Pkg folder can't find, Please check.")
                sys.exit()
            # Add find old version pkg.
            else:
                NeedProcOldPkg.append(temp[0])
    print("\nOld packages found:\n" + str(NeedProcOldPkg))

    #=================Create need process old/new Pkg info table===============================
    OldProcPkgInfo = []
    OldProcPkgInfo = [Proc.split("_") for Proc in NeedProcOldPkg] # Split to process
    NewProcPkgInfo = [Proc.split("_") for Proc in NeedProcOldPkg] # New Pkg Name List
    ProjectNameInfo = [Proc.split("_")[0] for Proc in NeedProcOldPkg] # ex:['Pacman', 'Asteroid']
    for OldProcPkg in OldProcPkgInfo:
        #======For Intel Project
        if Platform_Flag(OldProcPkg) in Intel_Platforms:
            # Save the info string to list
            for Proc in range(len(NewProcPkgInfo)):
                NewProcPkgInfo[Proc] = NewProcPkgInfo[Proc][:3] # ex:['Harp', 'MV', 'Q21']
                NewProcPkgInfo[Proc].append(NewVersion) # ex:['Harp', 'MV', 'Q21', 'NewVersion']
                if NewBuildID == "0000":
                    NewBuildID = ""
                if not (NewBuildID == "" or NewBuildID == "0000"):
                    NewProcPkgInfo[Proc].append(NewBuildID)
            break # Otherwise it will run more times
        #======For AMD Project
        else:
            # Save the info string to list
            NewVersion_AMD = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] # ex.['Q26', '02.07.03']
            for Proc in range(len(NeedProcOldPkg)):
                # For AMD special project setting.
                if Platform_Flag(Project) in AMD_Platforms_Special:
                    NewVersion_AMD = NewVersion # ex.'020703'
                    if len(NewProcPkgInfo[Proc]) == 3:
                        del NewProcPkgInfo[Proc][2]
                    NewProcPkgInfo[Proc][1] = NewVersion_AMD # ex:['Q26', 'NewVersion_AMD(02.07.03)']
                if (Platform_Flag(OldProcPkg) == "R24"):
                    NewProcPkgInfo[Proc] = NewProcPkgInfo[Proc][:2] # ex:['Worf', 'R24']
                    NewProcPkgInfo[Proc].append(NewVersion)
                    NewProcPkgInfo[Proc][2] = NewVersion_AMD # ex:['Worf', 'R24','NewVersion_AMD(02.07.03)']
            if NewBuildID == "0000":
                NewBuildID = ""
            if not (NewBuildID == "" or NewBuildID == "0000"):
                NewProcPkgInfo[Proc].append(NewBuildID)
            break # Otherwise it will run more times
    logging.debug("OldProcPkgInfo:" + str(OldProcPkgInfo))
    logging.debug("NewProcPkgInfo:" + str(NewProcPkgInfo))
    logging.debug("ProjectNameInfo:" + str(ProjectNameInfo))

    #=================Find Fv Folder Or Zip File=============================================
    print("Find Fv Folder Or Zip File".center(90, "="))
    # Look for the fv folder or zip file in the directory.
    Match_folder_list = FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)
    Match_zip_list = FindFvZip(ProcessProjectList, ProjectNameInfo, NewVersion, NewBuildID)

    for PkgInfo in range(len(NewProcPkgInfo)):
        #======For AMD Case Specialization Start
        # ex.['Q26', '01.04.01']=>['ScottyRr', 'ScottyRr', 'Q26', '010401']
        if Platform_Flag(OldProcPkgInfo) in AMD_Platforms_ExceptR24:
            NewProcPkgInfo[PkgInfo][1] = NewVersion
            NewProcPkgInfo[PkgInfo].insert(0, Config_data['AMDProjectName'][NewProcPkgInfo[PkgInfo][0]])
            NewProcPkgInfo[PkgInfo].insert(0, Config_data['AMDProjectName'][NewProcPkgInfo[PkgInfo][1]])
        # ex.['Worf', 'R24', '02.07.03']=>['Worf', 'Worf', 'R24', '020703']
        if Platform_Flag(OldProcPkgInfo) == "R24":
            NewProcPkgInfo[PkgInfo][2] = NewVersion
            NewProcPkgInfo[PkgInfo].insert(0, NewProcPkgInfo[PkgInfo][0]) # add "Worf"
        #======For AMD Case Specialization End

    print("Your Fv Folder: %s" % str(Match_folder_list))
    print("Your Fv Zip File: %s" % str(Match_zip_list))
    sys.exit()
    # If can't find Fv folder or Zip file.
    if len(Match_folder_list) == 0 and len(Match_zip_list) == 0:
        print("Can't find Fv folder and zip file.\nDownload Fv files from Production Release FTP.\n")
        temp = Ftp_multi(NewProcPkgInfo, Config_data['ProductionReleaseServer'], Config_data['TestReleaseServer'])[:]
        for name in temp:
            if str(name).find(".zip") != -1:
                Match_zip_list.append(name)
    # Number of Fv folders not match Project list
    elif len(Match_folder_list) < len(ProcessProjectList) and len(Match_folder_list) != 0 and len(Match_zip_list) < len(ProcessProjectList):
        print("Number of Fv folders not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp = Ftp_multi(NewProcPkgInfo, Config_data['ProductionReleaseServer'], Config_data['TestReleaseServer'])[:]
        for name in temp:
            if str(name).find(".zip") != -1:
                Match_zip_list.append(name)
    # Can't find Fv folders and Number of Fv Zip files not match Project list
    elif len(Match_folder_list) == 0 and len(Match_zip_list) < len(ProcessProjectList):
        print("Fv Zip files not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp = Ftp_multi(NewProcPkgInfo, Config_data['ProductionReleaseServer'], Config_data['TestReleaseServer'])[:]
        for name in temp:
            if str(name).find(".zip") != -1:
                Match_zip_list.append(name)
    # If Fv folder already exists
    elif len(Match_folder_list) == len(ProcessProjectList):
        print("Fv folder already exists.\n")

    Match_folder_list = FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)

    # Find Fv Zip file Start extracting
    if len(Match_zip_list) != -1:
        print("\nFind Fv Zip File, Start Extracting.")
        for i in range(len(Match_zip_list)):
            Foldername = Match_zip_list[i].replace(".zip", "")
            # For git hub package file
            if (str(Match_zip_list[i]).find(str(ProjectNameInfo[i].lower())) != -1) and (str(Match_zip_list[i]).find("Fv_") == -1) \
                and not os.path.isdir(".\\" + Foldername):
                UnZip(Match_zip_list[i])
                os.rename(".\\" + Foldername + "\\Fv", ".\\" + Foldername + "\\Fv_" + ProcessProjectList[i] + "_" + NewVersion + "_32")
                if not os.path.isdir(".\\Fv_" + ProcessProjectList[i] + "_" + NewVersion + "_32"):
                    move(".\\" + Foldername + "\\Fv_" + ProcessProjectList[i] + "_" + NewVersion + "_32", ".\\")
                rmtree(Foldername)
                # os.remove(".\\" + Match_zip_list[i]) #remove zip file
                print(Match_zip_list[i] + " Extract succeeded.")
                Match_folder_list.append("Fv_" + ProcessProjectList[i] + "_" + NewVersion + "_32")
            # For normal package file
            elif (str(Match_zip_list[i]).find("Fv_") != -1) and not os.path.isdir(".\\" + Foldername):
                UnZip(Match_zip_list[i])
                move(".\\" + Match_zip_list[i], ".\\" + Foldername)
                print(Match_zip_list[i] + " Extract succeeded.")
                Match_folder_list.append(Foldername)
            elif os.path.isdir(".\\" + Foldername):
                print("Fv folder " + Foldername + " already exists, Remove Fv zip file.")
                if os.path.isfile(".\\" + Match_zip_list[i]):
                    os.remove(".\\" + Match_zip_list[i])
        print("\nNow Your Fv Folder: %s" % str(str(Match_folder_list)))
    else:
        print("\nNow Your Fv Folder: %s" % str(Match_folder_list))

    # Working with multiple folders
    MatchMultipleFolder(Match_folder_list)

    #=================Find New Pkg Or Add New Pkg=============================================
    print("Find New Pkg Or Add New Pkg".center(90, "="))

    for PkgInfo in range(len(NewProcPkgInfo)):
        #======For AMD Start
        if Platform_Flag(OldProcPkgInfo) in AMD_Platforms_ExceptR24: # ['ScottyRr', 'ScottyRr', 'Q26', '010401']=>['Q26', '01.04.01']
            NewProcPkgInfo[PkgInfo].remove(Config_data['AMDProjectName'][NewProcPkgInfo[PkgInfo][2]])
            NewProcPkgInfo[PkgInfo].remove(Config_data['AMDProjectName'][NewProcPkgInfo[PkgInfo][1]])
            if Platform_Flag(OldProcPkgInfo) == "Q26" or Platform_Flag(OldProcPkgInfo) == "Q27":
                NewProcPkgInfo[PkgInfo][1] = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
        if Platform_Flag(OldProcPkgInfo) == "R24": # ['Worf', 'Worf', 'R24, '010401']=>['Worf', 'R24', '01.04.01']
            NewProcPkgInfo[PkgInfo].remove(NewProcPkgInfo[PkgInfo][0]) # remove 'Worf'
            NewProcPkgInfo[PkgInfo][1] = Platform_Flag(OldProcPkgInfo)
            NewProcPkgInfo[PkgInfo][2] = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
        #======For AMD End

    for OProc in range(len(OldProcPkgInfo)): # How much Old Version folder
        OldVersionPath = ".\\" + ("_").join(OldProcPkgInfo[OProc])
        NewVersionPath = ".\\" + ("_").join(NewProcPkgInfo[OProc])
        if not os.path.isdir(NewVersionPath): # Check NewVersion Folder Exist
            if not os.path.isdir(OldVersionPath + "\\" + ("_").join(OldProcPkgInfo[OProc])): # Check Old Pkg is in folder???
                Copy_Release_Folder(OldVersionPath, NewVersionPath)
            elif not os.path.isdir(OldVersionPath + "\\FPTW"):
                if os.path.isdir(OldVersionPath + "\\" + ("_").join(OldProcPkgInfo[OProc])): # Check Old Pkg is in folder???
                    Copy_Release_Folder(OldVersionPath + "\\" + ("_").join(OldProcPkgInfo[OProc]), NewVersionPath)
                else:
                    print("Pkg " + ("_").join(OldProcPkgInfo[OProc]) + " can't find.")
            else:
                print("Pkg " + ("_").join(OldProcPkgInfo[OProc]) + " can't find.")
        else:
            print("Pkg " + NewVersionPath.split("\\")[-1] + Fore.RED + " already exists.")
            sys.exit()
    if len(Match_folder_list) == 0:
        print("Can't find anything Fv folder.\n")

    #=================Modify Pkg Update Version==============================================
    print("Modify Pkg Update Version".center(90, "="))
    BiosBuildDate = CheckBiosBuildDate(Match_folder_list)
    BiosMrcVersion = GetMrcVersion(Match_folder_list)
    BiosIshVersion = GetIshVersion(Match_folder_list)
    BiosPmcVersion = GetPmcVersion(Match_folder_list)
    BiosNphyVersion = GetNphyVersion(Match_folder_list)
    BiosBinaryChecksum = CheckFileChecksum(Match_folder_list, NewVersion)
    for NProc in NewProcPkgInfo: # Pkg Modify Update Version
        #======For Intel Project
        if Platform_Flag(NProc) in Intel_Platforms:
            Path = os.getcwd() + "\\" + ("_").join(NProc)
            if os.path.isdir(Path+"\\FPTW"): # Check Folder Exist
                ReleaseNote_docx = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("Release" in ReleaseNote) and (".docx" in ReleaseNote)]
                ReleaseNote_xlsm = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("Release" in ReleaseNote) and ("Note" in ReleaseNote) and (".xlsm" in ReleaseNote)]
                if len(ReleaseNote_docx) == 1: # If get release note G4
                    os.chdir(Path)
                    os.rename(ReleaseNote_docx[0], ("_").join(NProc) + " release note.docx")
                    if os.path.isfile(("_").join(NProc)+" release note.docx"):
                        print("ReleaseNote Rename to " + ("_").join(NProc) + " release note.docx" + " succeeded.")
                elif len(ReleaseNote_xlsm) == 1: # If get release note G5 and late
                    os.chdir(Path)
                    ReleaseName = NProc[0] + "_" + NProc[1] + "_" + NProc[2] + "_BIOS_Release_Note_"
                    if (NewBuildID == "" or NewBuildID == "0000"):
                        if ReleaseNote_xlsm[0] == ReleaseName + NewVersion + ".xlsm":
                            print("ReleaseNote Already Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + NewVersion + ".xlsm")
                        if os.path.isfile(ReleaseName + NewVersion + ".xlsm"):
                            ReleaseNoteName = ReleaseName + NewVersion + ".xlsm"
                            ModifyReleaseNote(NProc, ReleaseNoteName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
                            print("ReleaseNote Rename to " + ReleaseName + NewVersion + ".xlsm" + " succeeded.")
                    else:
                        if ReleaseNote_xlsm[0] == ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm":
                            print("ReleaseNote Already Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm")
                        if os.path.isfile(ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm"):
                            ReleaseNoteName = ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm"
                            ModifyReleaseNote(NProc, ReleaseNoteName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
                            print("ReleaseNote Rename to " + ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm" + " succeeded.")
                else:
                    print("Can't find release note file.")
                os.chdir(Path + "\\FPTW")
                ChangeBuildID(NProc, Config_data['VersionFileList'], NewVersion)
            else:
                print("Pkg Folder " + ("_").join(NProc) + " can't find.\n")
            os.chdir("..\..")
        #======For AMD Project
        else:
            Path = os.getcwd() + "\\" + ("_").join(NProc)
            if os.path.isdir(Path + "\\AMDFLASH"): # Check Folder Exist
                ReleaseNote_docx = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("elease" in ReleaseNote) and (".docx" in ReleaseNote)]
                ReleaseNote_xlsm = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("Release" in ReleaseNote) and ("Note" in ReleaseNote) and (".xlsm" in ReleaseNote)]
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
                            print("ReleaseNote Alreadly Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + "Note_" + NewVersion + ".xlsm")
                        if os.path.isfile(ReleaseName + "Note_" + NewVersion + ".xlsm"):
                            ReleaseNoteName = ReleaseName + "Note_" + NewVersion + ".xlsm"
                            ModifyReleaseNote(NProc, ReleaseNoteName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
                            print("ReleaseNote Rename to " + ReleaseName + "Note_" + NewVersion + ".xlsm" + " succeeded.")
                    else:
                        if ReleaseNote_xlsm[0] == ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm":
                            print("ReleaseNote Alreadly Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm")
                        if os.path.isfile(ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm"):
                            ReleaseNoteName = ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm"
                            ModifyReleaseNote(NProc, ReleaseNoteName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, BiosIshVersion, BiosPmcVersion, BiosNphyVersion, Match_folder_list)
                            print("ReleaseNote Rename to " + ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm" + " succeeded.")
                else:
                    print("Can't find Release_Notes.docx")
                os.chdir(Path + "\\AMDFLASH")
                ChangeBuildID(NProc, Config_data['VersionFileList'], NewVersion)
            else:
                print("Pkg " + ("_").join(NProc) + " can't find.\n")
            os.chdir("..\..")

    #=================Remove Pkg Old File=====================================================
    print("Remove Pkg Old File".center(90, "="))
    for NProc in NewProcPkgInfo:
        target_folder = ("_").join(NProc)
        if os.path.isdir(".\\"+target_folder):
            RemoveOldFileInDir(target_folder, Config_data['RemoveFileRule'], Config_data['NotRemoveFileRule'])
        else:
            print("Pkg "+("_").join(NProc)+" can't find.\n")

    #=================Fv File Rename And Copy To Pkg===========================================
    print("Fv File Rename And Copy To Pkg".center(90, "="))
    for Fv in Match_folder_list:
        for NProc in NewProcPkgInfo:
            #======For Intel Project
            if Platform_Flag(NProc) in Intel_Platforms:
                if Fv.split("_")[1] == NProc[2]:
                    if os.path.isdir(".\\"+Fv):
                        Path = ".\\" + Fv
                        Board_version = NProc[2]+"_" + NProc[3]
                        if os.path.isfile(Path + "\\" + Board_version + "_12.bin") or os.path.isfile(Path + "\\" + Board_version + "_16.bin"): # If Alreadly renamed
                            if os.path.isfile(Path+"\\"+Board_version+".xml"):
                                print(Board_version + "_12.bin or _16.bin & " + Board_version + ".xml alreadly renamed.")
                        if os.path.isfile(Path + "\\" + Board_version + ".bin") and os.path.isfile(Path + "\\" + Board_version + "_16.bin"):
                            os.rename(Path + "\\" + Board_version + ".bin", Path + "\\" + Board_version + "_12.bin") # Rename Fv folder 2 files
                            os.rename(Path + "\\" + NProc[2] + ".xml", Path + "\\" + Board_version + ".xml")
                            print(Board_version + "_12.bin & " + Board_version + ".xml rename succeeded.")
                        if os.path.isfile(Path + "\\" + Board_version + ".bin") and os.path.isfile(Path + "\\" + Board_version + "_32.bin"):
                            os.rename(Path + "\\" + Board_version + ".bin", Path + "\\" + Board_version + "_16.bin") # Rename Fv folder 2 files
                            os.rename(Path + "\\" + NProc[2] + ".xml", Path + "\\" + Board_version + ".xml")
                            print(Board_version + "_16.bin & " + Board_version + ".xml rename succeeded.")
                        if (os.path.isfile(Path+"\\"+Board_version+"_32.bin") or os.path.isfile(Path+"\\"+Board_version+"_16.bin")) \
                            and os.path.isdir(".\\"+("_").join(NProc)): # Check Pkg Folder Exist
                            Copy_Release_Files(Fv, ("_").join(NProc), NProc, Match_folder_list)
                        else:
                            print("Pkg " + ("_").join(NProc) + " can't find.")
                    else:
                        print("Need to be processed Fv folder:" + Fv + " can't find.\n")
            #======For AMD Project
            else:
                if Platform_Flag(NProc) in AMD_Platforms:
                    if (Fv.split("_")[1] == NProc[0]) or (Fv.split("_")[1] == NProc[1]):
                        if os.path.isdir(".\\" + Fv):
                            Path = os.getcwd() + "\\" + Fv
                            Board_version = NProc[0] + "_" + NewVersion
                        if Platform_Flag(NProc) == "R24":
                            Board_version = NProc[1] + "_" + NewVersion
                        if os.path.isfile(Path + "\\" + Board_version + ".bin") or os.path.isfile(Path + "\\" + Board_version + "_16.bin") or os.path.isfile(Path + "\\" + Board_version + "_32.bin"): # For 16MB BIOS
                            if os.path.isfile(Path + "\\" + Board_version[:3] + ".xml"):
                                if os.path.isdir(".\\" + ("_").join(NProc)): # Check Pkg Folder Exist
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
            #======For Intel G5&G6 and late Project
            if Platform_Flag(NProc) in Intel_Platforms and os.path.isdir(".\\" + ("_").join(NProc) + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)"):
                Tool_version_table_path = ".\\" + ("_").join(NProc) + "\\FactoryUtility\\ToolVersion.xlsx"
                Tool_version_info = ReadToolVersionTable(Tool_version_table_path)
                Check = "Match"
                for name, type, verinfo, path, note in Tool_version_info:
                    ver = ChangeVersionInfo(verinfo)
                    #date = ChangeDataInfo(dateinfo)
                    CompareInfo(NProc, name, ver, path, Tool_version_table_path)
            #======For ADM G5 and late Project
            elif Platform_Flag(NProc) in AMD_Platforms_Special and os.path.isdir(".\\" + ("_").join(NProc) + "\\Capsule\\Windows"):
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
    if Platform_Flag(OldProcPkgInfo) in Intel_Platforms:
        CheckPkg(NewProcPkgInfo) # Check new release Pkg is OK?
        PrintBiosBuildDate(Match_folder_list, BiosBuildDate)
        PrintBiosBinaryChecksum(NewProcPkgInfo, BiosBinaryChecksum, NewVersion)
    #======For AMD Check
    else:
        CheckPkg_AMD(NewProcPkgInfo, NewVersion, NewBuildID) # Check new release Pkg is OK?
        PrintBiosBuildDate(Match_folder_list, BiosBuildDate)
        PrintBiosBinaryChecksum(NewProcPkgInfo, BiosBinaryChecksum, NewVersion)
    print("\nFinally pkg please compare with leading project.\n")
    os.system('Pause')
    sys.exit()
    # =================Script End===============================================================