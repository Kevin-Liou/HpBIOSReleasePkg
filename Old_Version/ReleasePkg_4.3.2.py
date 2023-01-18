#=================================
# coding=UTF-8
# Make Release Pkg Script.   4.3.2
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com

#This script is for makeing release pkg
#Now can use Intel DM G4, G5, AMD DM G4 platform
#=================================

from Lib import *

# In add the Projects to be here, and the following BoradID should also add.
BoardID=["P10","Q10","Q11","Q21","Q22","Q23","Q26","Q27","R11","R21","R22","R23","R26"]
ProjectName={"Q26":"ScottyRr", "Q27":"Scotty", "R26":"Riker"}

Versionfilelist=["BUFF2.nsh", "Buff2All.nsh", "Update32.bat", "Update64.bat", "UpdateEFI.nsh",
                 "Update32_vPro.bat", "Update64_vPro.bat", "UpdateEFI_vPro.nsh"]

NotRemovefilerule=["Note", "note", "History", "How to Flash"] # Priority over "Removefilerule"

Removefilerule=["DCI.7z", ".cer", ".pfx", ".pvk", r"\d{6}.cat", r"\w\d{2}_\d{4}.cat", r"P00\w{3}-\w{3}.zip",
                r"\d\d_\d\d_\d\d.bin", r"\d\d_\d\d_\d\d.cat", r"\d\d_\d\d_\d\d.inf", r"Build.Log",
                r"\w\d{2}_\d{6}.xml", r"\w\d{2}_\d{4}.xml", ".xlsm",
                r"\d{4}_12.bin", r"\d{4}_16.bin", r"\d{4}_32.bin",
                r"\d{6}.bin", r"\d{6}.inf", r"\w\d{2}_\d{4}.bin", r"\w\d{2}_\d{4}.inf"]
                #r"ME_+[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]+.bin"] For Intel ME

ProductionReleaseServer={   "host":"ftp.usa.hp.com",
                            "username":"sign_ron",
                            "pwd":"5gg9*0UA"}

TestReleaseServer={         "host":"ftp.usa.hp.com",
                            "username":"bios15ws",
                            "pwd":"e.QV5ra}"}

    # Script Start
if __name__ == '__main__':
    init(autoreset=True)# colorama
    freeze_support()# For windows do multiprocessing.

    #=================Script Start==========================================================
    print("Make Release Pkg Script   4.3.2".center(90, "="))
    print("Input Information".center(90, "="))
    print("This script is for makeing release pkg.\nNow can use Intel DM G4, G5, AMD DM G4 platform.\n")
    OldVersion=InputStr("OldVersion:")# Input OldVersion & NewVersion & NewBuildID
    NewVersion=InputStr("NewVersion:")
    NewBuildID=InputStr("NewBuildID:")# ex: 020106_"0001"
    ProcessProject=InputStr("                                   (Can Multiple Select)\
        \nPlease Enter Projects To Processed"+str(BoardID)+":")# Input need Process boardID, Can multiple choice
    ProcessProjectList=ProcessProject.upper().split()
    if OldVersion=="":
        print("\nPlease Input OldVersion.")
        sys.exit()
    elif NewVersion=="":
        print("\nPlease Select NewVersion.")
        sys.exit()

    #=================Find Need Process Old Pkg==============================================
    print("Find Need Process Old Pkg".center(90, "="))
    OldBuildID="0";NeedProcOldPkg=[]
    print("Start Find Process Project Old Pkg.")
    for Project in ProcessProjectList:
        temp=[]
        #======For Intel
        if Platform_Flag(Project) == "Intel" or Platform_Flag(Project) == "Intel G3":
            for Dir in os.listdir(".\\"):# Find oldveriosn pkg.
                if not Dir.split("_")[0] == "Fv" and not Dir.find(".7z") != -1 and not Dir.find(".zip") !=- 1:
                    if Project+"_"+OldVersion in Dir:
                        temp.append(Dir)
            if len(temp) > 1 and OldBuildID == "0":# If find oldveriosn pkg have buildID.
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
            if len(temp) == 0:# Can't find oldveriosn pkg.
                print(Project + "_" + OldVersion + " Old Pkg folder can't find, Please check.")
            else:# Add find oldveriosn pkg.
                NeedProcOldPkg.append(temp[0])
        #======For AMD
        else:
            OldVersion_AMD = OldVersion[0:2]+"."+OldVersion[2:4]+"."+OldVersion[4:6] # ex.['Q26', '02.07.03']
            if Platform_Flag(Project) == "R26":# 78 AMD 78 R26
                OldVersion_AMD = OldVersion
            for Dir in os.listdir(".\\"):# Find oldveriosn pkg.
                if not Dir.split("_")[0]=="Fv" and not Dir.find(".7z")!=-1 and not Dir.find(".zip")!=-1:
                    if Project+"_"+OldVersion_AMD in Dir:
                        temp.append(Dir)
            if len(temp)>1 and OldBuildID=="0":# If find oldveriosn pkg have buildID.
                OldBuildID=InputStr("\n"+str(temp)+"\n"+Project+"_"+OldVersion_AMD+" Please Select OldBuildID:")
                if OldBuildID=="" or OldBuildID=="0000":
                    temp=[a for a in temp if len(a.split("_"))==2]
                else:
                    temp=[a for a in temp if Project+"_"+OldVersion_AMD+"_"+OldBuildID in a]
                if not len(temp)==1:
                    temp=[Dir for Dir in os.listdir(".\\") if Project+"_"+OldVersion_AMD in Dir]
                    OldBuildID=InputStr("\n"+str(temp)+"\n"+Project+"_"+OldVersion_AMD+" Please Select OldBuildID:")
                    if OldBuildID=="" or OldBuildID=="0000":
                        temp=[a for a in temp if len(a.split("_"))==2]
                    else:
                        temp=[a for a in temp if Project+"_"+OldVersion_AMD+"_"+OldBuildID in a]
            if len(temp)==0:# Can't find oldveriosn pkg.
                print(Project+"_"+OldVersion_AMD+" Old Pkg folder can't find, Please check.")
            else:# Add find oldveriosn pkg.
                NeedProcOldPkg.append(temp[0])
    print("\nYour need process old Pkg:\n"+str(NeedProcOldPkg))

    #=================Make need process old/new Pkg info table===============================
    OldProcPkgInfo = []
    OldProcPkgInfo = [Proc.split("_") for Proc in NeedProcOldPkg]# Split to process
    NewProcPkgInfo = [Proc.split("_") for Proc in NeedProcOldPkg]# New Pkg Name List
    for OldProcPkg in OldProcPkgInfo:
        #======For Intel
        if Platform_Flag(OldProcPkg) == "Intel" or Platform_Flag(OldProcPkg) == "Intel G3":
            for Proc in range(len(NewProcPkgInfo)):
                NewProcPkgInfo[Proc]=NewProcPkgInfo[Proc][:3] # ex:['Harp', 'MV', 'Q21']
                NewProcPkgInfo[Proc].append(NewVersion) # ex:['Harp', 'MV', 'Q21', 'NewVersion']
                if NewBuildID == "0000":
                    NewBuildID = ""
                if not (NewBuildID == "" or NewBuildID == "0000"):
                    NewProcPkgInfo[Proc].append(NewBuildID)
            break #Otherwise it will run more times
        #======For AMD
        else:
            NewVersion_AMD = NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6] # ex.['Q26', '02.07.03']
            if Platform_Flag(OldProcPkg) == "R26":# 78 AMD 78 R26
                NewVersion_AMD = NewVersion
            for Proc in range(len(NeedProcOldPkg)):
                if len(NewProcPkgInfo[Proc]) == 3:
                    del NewProcPkgInfo[Proc][2]
                NewProcPkgInfo[Proc][1] = NewVersion_AMD # ex:['Q26', 'NewVersion_AMD']
                if NewBuildID == "0000":
                    NewBuildID = ""
                if not (NewBuildID == "" or NewBuildID == "0000"):
                    NewProcPkgInfo[Proc].append(NewBuildID)
            break #Otherwise it will run more times

    #=================Find Fv Folder Or Zip File=============================================
    print("Find Fv Folder Or Zip File".center(90, "="))
    Matchfolderlist = FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)
    Matchziplist = FindFvZip(ProcessProjectList, NewVersion, NewBuildID)

    for PkgInfo in range(len(NewProcPkgInfo)):
        #======For AMD Start # ['Q26', '01.04.01']=>['ScottyRr', 'ScottyRr', 'Q26', '010401']
        if Platform_Flag(OldProcPkgInfo) == "AMD" or Platform_Flag(OldProcPkgInfo) == "R26":
            NewProcPkgInfo[PkgInfo][1] = NewVersion
            NewProcPkgInfo[PkgInfo].insert(0, ProjectName[NewProcPkgInfo[PkgInfo][0]])
            NewProcPkgInfo[PkgInfo].insert(0, ProjectName[NewProcPkgInfo[PkgInfo][1]])
        #======For AMD End

    print("Your Fv Folder: %s" % str(Matchfolderlist))
    print("Your Fv Zip File: %s" % str(Matchziplist))
    # If can't find Fv folder or Zip file.
    if len(Matchfolderlist) == 0 and len(Matchziplist) == 0:
        print("Can't find Fv folder and zip file.\nDownload Fv files from Production Release FTP.\n")
        temp=Ftp_multi(NewProcPkgInfo, ProductionReleaseServer, TestReleaseServer)[:]
        for name in temp:
            if str(name).find(".zip") != -1:
                Matchziplist.append(name)
    # Number of Fv folders not match Projectlist
    elif len(Matchfolderlist) < len(ProcessProjectList) and len(Matchfolderlist) != 0 and len(Matchziplist) < len(ProcessProjectList):
        print("Number of Fv folders not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp=Ftp_multi(NewProcPkgInfo, ProductionReleaseServer, TestReleaseServer)[:]
        for name in temp:
            if str(name).find(".zip") != -1:
                Matchziplist.append(name)
    # Can't find Fv folders and Number of Fv Zip files not match Projectlist
    elif len(Matchfolderlist) == 0 and len(Matchziplist)<len(ProcessProjectList):
        print("Fv Zip files not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp=Ftp_multi(NewProcPkgInfo, ProductionReleaseServer, TestReleaseServer)[:]
        for name in temp:
            if str(name).find(".zip") != -1:
                Matchziplist.append(name)
    # If Fv folder already exists
    elif len(Matchfolderlist) == len(ProcessProjectList):
        print("Fv folder already exists.\n")
    
    Matchfolderlist = FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)
    
    # Find Fv Zip file Start extracting
    if len(Matchziplist) != -1:
        print("\nFind Fv Zip File, Start Extracting.")
        for Zipname in Matchziplist:
            if not os.path.isdir(".\\"+Zipname.split(".")[0]):
                UnZip(Zipname)
                move(".\\"+Zipname, ".\\"+Zipname.split(".")[0])
                print(Zipname+" Extract succeeded.")
                Matchfolderlist.append(Zipname.split(".")[0])
            else:
                print("Fv folder "+Zipname.split(".")[0]+" already exists, Remove Fv zip file.")
                os.remove(".\\"+Zipname)
        print("\nNow Your Fv Folder: %s" % str(str(Matchfolderlist)))
    else:
        print("\nNow Your Fv Folder: %s" % str(Matchfolderlist))

    #=================Find New Pkg Or Add New Pkg=============================================
    print("Find New Pkg Or Add New Pkg".center(90, "="))

    for PkgInfo in range(len(NewProcPkgInfo)):
        #======For AMD Start # ['ScottyRr', 'ScottyRr', 'Q26', '010401']=>['Q26', '01.04.01']
        if Platform_Flag(OldProcPkgInfo) == "AMD" or Platform_Flag(OldProcPkgInfo) == "R26":
            NewProcPkgInfo[PkgInfo].remove(ProjectName[NewProcPkgInfo[PkgInfo][2]])
            NewProcPkgInfo[PkgInfo].remove(ProjectName[NewProcPkgInfo[PkgInfo][1]])
            if Platform_Flag(OldProcPkgInfo) == "AMD":
                NewProcPkgInfo[PkgInfo][1] = NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]
        #======For AMD End

    for OProc in range(len(OldProcPkgInfo)):# How much Old Version folder
        OldVersionPath = ".\\" + ("_").join(OldProcPkgInfo[OProc])
        NewVersionPath = ".\\" + ("_").join(NewProcPkgInfo[OProc])
        if not os.path.isdir(NewVersionPath):# Check NewVersion Folder Exist
            if not os.path.isdir(OldVersionPath+"\\"+("_").join(OldProcPkgInfo[OProc])):# Check Old Pkg is in folder???
                Copy_Release_Folder(OldVersionPath, NewVersionPath)
            elif not os.path.isdir(OldVersionPath+"\\FPTW"):
                if os.path.isdir(OldVersionPath+"\\"+("_").join(OldProcPkgInfo[OProc])):# Check Old Pkg is in folder???
                    Copy_Release_Folder(OldVersionPath+"\\"+("_").join(OldProcPkgInfo[OProc]), NewVersionPath)
                else:
                    print("Pkg "+("_").join(OldProcPkgInfo[OProc])+" can't find.")
            else:
                print("Pkg "+("_").join(OldProcPkgInfo[OProc])+" can't find.")
        else:
            print("Pkg "+NewVersionPath.split("\\")[-1]+Fore.RED+" already exists.")
            sys.exit()
    if len(Matchfolderlist) == 0: print("Can't find anything Fv folder.\n")

    #=================Modify Pkg Update Version==============================================
    print("Modify Pkg Update Version".center(90, "="))
    BiosBuildDate = CheckBiosBuildDate(Matchfolderlist)
    BiosBinaryCRC32 = CheckFileCRC32(Matchfolderlist, NewVersion, OldVersion)
    for NProc in NewProcPkgInfo:# Pkg Modify Update Version
        #======For Intel
        if Platform_Flag(NProc) == "Intel" or Platform_Flag(NProc) == "Intel G3":
            Path=os.getcwd()+"\\"+("_").join(NProc)
            if os.path.isdir(Path+"\\FPTW"):# Check Folder Exist
                ReleaseNote_docx = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("elease" in ReleaseNote) and (".docx" in ReleaseNote)]
                ReleaseNote_xlsm = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("Release_Note" in ReleaseNote) and (".xlsm" in ReleaseNote)]
                if len(ReleaseNote_docx) == 1:
                    os.chdir(Path)
                    os.rename(ReleaseNote_docx[0], ("_").join(NProc)+" release note.docx")
                    if os.path.isfile(("_").join(NProc)+" release note.docx"):
                        print("ReleaseNote Rename to "+("_").join(NProc)+" release note.docx"+" succeeded.")
                elif len(ReleaseNote_xlsm) == 1:
                    os.chdir(Path)
                    ReleaseName = NProc[0]+"_"+NProc[1]+"_"+NProc[2]+"_BIOS_Release_Note_"
                    if (NewBuildID == "" or NewBuildID == "0000"):
                        if ReleaseNote_xlsm[0] == ReleaseName + NewVersion + ".xlsm":
                            print("ReleaseNote Alreadly Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + NewVersion + ".xlsm")
                        if os.path.isfile(ReleaseName + NewVersion + ".xlsm"):
                            ModifyReleaseNoteForG5(NProc, ReleaseName + NewVersion + ".xlsm", BiosBuildDate, BiosBinaryCRC32, NewVersion, NewBuildID)
                            print("ReleaseNote Rename to " + ReleaseName + NewVersion + ".xlsm" + " succeeded.")
                    else:
                        if ReleaseNote_xlsm[0] == ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm":
                            print("ReleaseNote Alreadly Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm")
                        if os.path.isfile(ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm"):
                            ModifyReleaseNoteForG5(NProc, ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm", BiosBuildDate, BiosBinaryCRC32, NewVersion, NewBuildID)
                            print("ReleaseNote Rename to " + ReleaseName + NewVersion + "_" + NewBuildID + ".xlsm" + " succeeded.")
                else:
                    print("Can't find release note file.")
                os.chdir(Path+"\\FPTW")
                ChangeBuildID(NProc, Versionfilelist, NewVersion)
            else:
                print("Pkg Folder "+("_").join(NProc)+" can't find.\n")
            os.chdir("..\..")
        #======For AMD
        else:
            Path=os.getcwd()+"\\"+("_").join(NProc)
            if os.path.isdir(Path+"\\AMDFLASH"):# Check Folder Exist
                ReleaseNote_docx = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("elease" in ReleaseNote) and (".docx" in ReleaseNote)]
                ReleaseNote_xlsm = [ReleaseNote for ReleaseNote in os.listdir(Path) if ("Release_Note" in ReleaseNote) and (".xlsm" in ReleaseNote)]
                if len(ReleaseNote_docx) == 1:
                    os.chdir(Path)
                    os.rename(ReleaseNote_docx[0], "Scotty_"+("_").join(NProc)+"_Release_Notes.docx")
                    if os.path.isfile("Scotty_"+("_").join(NProc)+"_Release_Notes.docx"):
                        print("ReleaseNote Rename to "+"Scotty_"+("_").join(NProc)+"_Release_Notes.docx"+" succeeded.")
                elif len(ReleaseNote_xlsm) == 1:
                    os.chdir(Path)
                    ReleaseName = ReleaseNote_xlsm[0].split("Note_")[0]
                    if (NewBuildID == "" or NewBuildID == "0000"):
                        if ReleaseNote_xlsm[0] == ReleaseName + "Note_" + NewVersion + ".xlsm":
                            print("ReleaseNote Alreadly Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + "Note_" + NewVersion + ".xlsm")
                        if os.path.isfile(ReleaseName + "Note_" + NewVersion + ".xlsm"):
                            ModifyReleaseNoteForG5(NProc, ReleaseName + "Note_" + NewVersion + ".xlsm", BiosBuildDate, BiosBinaryCRC32, NewVersion, NewBuildID)
                            print("ReleaseNote Rename to " + ReleaseName + "Note_" + NewVersion + ".xlsm" + " succeeded.")
                    else:
                        if ReleaseNote_xlsm[0] == ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm":
                            print("ReleaseNote Alreadly Rename.")
                            break
                        os.rename(ReleaseNote_xlsm[0], ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm")
                        if os.path.isfile(ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm"):
                            ModifyReleaseNoteForG5(NProc, ReleaseName+"Note_"+NewVersion+"_"+NewBuildID+".xlsm", BiosBuildDate, BiosBinaryCRC32, NewVersion, NewBuildID)
                            print("ReleaseNote Rename to " + ReleaseName + "Note_" + NewVersion + "_" + NewBuildID + ".xlsm" + " succeeded.")
                else:
                    print("Can't find Release_Notes.docx")
                os.chdir(Path+"\\AMDFLASH")
                ChangeBuildID(NProc, Versionfilelist, NewVersion)
            else:
                print("Pkg "+("_").join(NProc)+" can't find.\n")
            os.chdir("..\..")

    #=================Remove Pkg Old File=====================================================
    print("Remove Pkg Old File".center(90, "="))
    for NProc in NewProcPkgInfo:
        targetfolder=("_").join(NProc)
        if os.path.isdir(".\\"+targetfolder):
            RemoveOldFileInDir(targetfolder, Removefilerule, NotRemovefilerule)
        else:
            print("Pkg "+("_").join(NProc)+" can't find.\n")

    #=================Fv File Rename And Copy To Pkg===========================================
    print("Fv File Rename And Copy To Pkg".center(90, "="))
    for Fv in Matchfolderlist:
        for NProc in NewProcPkgInfo:
            #======For Intel
            if Platform_Flag(NProc) == "Intel" or Platform_Flag(NProc) == "Intel G3":
                if Fv.split("_")[1]==NProc[2]:
                    if os.path.isdir(".\\"+Fv):
                        Path = ".\\" + Fv
                        Boardversion = NProc[2]+"_" + NProc[3]
                        if os.path.isfile(Path + "\\" + Boardversion + "_12.bin") or os.path.isfile(Path + "\\" + Boardversion + "_16.bin"):# If Alreadly renamed
                            if os.path.isfile(Path+"\\"+Boardversion+".xml"):
                                print(Boardversion + "_12.bin or _16.bin & " + Boardversion + ".xml alreadly renamed.")
                        if os.path.isfile(Path + "\\" + Boardversion + ".bin") and os.path.isfile(Path + "\\" + Boardversion + "_16.bin"):
                            os.rename(Path + "\\" + Boardversion + ".bin", Path + "\\" + Boardversion + "_12.bin")  # Rename Fv folder 2 files
                            os.rename(Path + "\\" + NProc[2] + ".xml", Path + "\\" + Boardversion + ".xml")
                            print(Boardversion + "_12.bin & " + Boardversion + ".xml rename succeeded.")
                        if os.path.isfile(Path + "\\" + Boardversion + ".bin") and os.path.isfile(Path + "\\" + Boardversion + "_32.bin"):
                            os.rename(Path + "\\" + Boardversion + ".bin", Path + "\\" + Boardversion + "_16.bin")  # Rename Fv folder 2 files
                            os.rename(Path + "\\" + NProc[2] + ".xml", Path + "\\" + Boardversion + ".xml")
                            print(Boardversion + "_16.bin & " + Boardversion + ".xml rename succeeded.")
                        if os.path.isfile(Path+"\\"+Boardversion+".inf") and os.path.isdir(".\\"+("_").join(NProc)):# Check Pkg Folder Exist
                            Copy_Release_Files(Fv, ("_").join(NProc))
                        else:
                            print("Pkg "+("_").join(NProc)+" can't find.")
                    else:
                        print("Need to be processed Fv folder:"+Fv+" can't find.\n")
            #======For AMD
            else:
                if Fv.split("_")[1]==NProc[0]:
                    if os.path.isdir(".\\"+Fv):
                        Path = os.getcwd()+"\\" + Fv
                        Boardversion = NProc[0] + "_" + NewVersion
                        if os.path.isfile(Path+"\\"+Boardversion+".bin") or os.path.isfile(Path+"\\"+Boardversion+"_16.bin"):# For 16MB BIOS
                            if os.path.isfile(Path+"\\"+Boardversion+".inf") and os.path.isfile(Path+"\\"+NProc[0]+".xml"):
                                if os.path.isdir(".\\"+("_").join(NProc)):# Check Pkg Folder Exist
                                    Copy_Release_Files_AMD(Fv, ("_").join(NProc), NewVersion)
                                else:
                                    print("Pkg "+("_").join(NProc)+" can't find.")
                    else:
                        print("Need to be processed Fv folder:"+Fv+" can't find.\n")
    if len(Matchfolderlist) == 0 or len(ProcessProjectList) == 0:
        print("Can't find anything Fv folder.\n")

    #=================Check Tool Version Is Match In Table======================================
    print("Check Tool Version Is Match In Table".center(90, "=")+"\n")
    try:
        for NProc in NewProcPkgInfo:
            #======For Intel G5
            if Platform_Flag(NProc) == "Intel" and os.path.isdir(".\\"+("_").join(NProc)+"\\Capsule\\CCG5\\CCG5C"):
                Toolversiontablepath=".\\"+("_").join(NProc)+"\\FactoryUtility\\ToolVersion.xlsx"
                Toolversioninfo=ReadToolVersionTable(Toolversiontablepath)
                Check="Match"
                print("Name".ljust(28)+"File".ljust(16)+"Table".ljust(20)+"File".ljust(13)+"Table")
                for name, type, verinfo, path, none, dateinfo in Toolversioninfo:
                    ver=ChangeVersionInfo(verinfo)
                    date=ChangeDataInfo(dateinfo)
                    fileinfo_temp=GetFileInfo(".\\"+("_").join(NProc)+path+name)
                    if name=="EEUPDATEW32.exe" or name=="EEUPDATEW64e.exe":
                        if os.listdir(".\\"+("_").join(NProc)+"\\Global\\BIOS")[0].split("_")[2]=="16.bin":# If 400 project
                            continue
                        if fileinfo_temp[0]!=ver or fileinfo_temp[1]!=date:
                            print(name.ljust(28)+fileinfo_temp[0]+"==>"+str(ver)+"\t\t"+fileinfo_temp[1]+"==>"+str(date))
                            Check="Not Match"
                    elif name=="FWUpdLcl64.exe" or name=="ElectronicLabelUpdate.efi" or name=="wmiTool64.exe":# Is zip file
                        if name=="FWUpdLcl64.exe":
                            filedateinfo = GetZipDateInfo(".\\"+("_").join(NProc)+"\\FactoryUtility\\FWUpdate.zip", name)
                            if filedateinfo!=date:
                                print(name.ljust(64)+str(filedateinfo)+"==>"+str(date))
                                Check="Not Match"
                        if name=="ElectronicLabelUpdate.efi":
                            filedateinfo = GetZipDateInfo(".\\"+("_").join(NProc)+"\\FactoryUtility\\ElectronicLabelUpdate.zip", name)
                            if filedateinfo!=date:
                                print(name.ljust(64)+str(filedateinfo)+"==>"+str(date))
                                Check="Not Match"
                        if name=="wmiTool64.exe":
                            filedateinfo = GetZipDateInfo(".\\"+("_").join(NProc)+"\\FactoryUtility\\wmiTool.zip", name)
                            if filedateinfo!=date:
                                print(name.ljust(64)+str(filedateinfo)+"==>"+str(date))
                                Check="Not Match"
                    elif name=="Buff2.efi" or name=="EnableHAPFull.efi":# Is efi file
                        if fileinfo_temp[1]!=date:
                            print(name.ljust(64)+fileinfo_temp[1]+"==>"+str(date))
                            Check="Not Match"
                    elif name==None:
                        pass
                    else:
                        if fileinfo_temp[0]!=ver or fileinfo_temp[1]!=date:
                            print(name.ljust(28)+fileinfo_temp[0]+"==>"+ver+"\t\t"+fileinfo_temp[1]+"==>"+date)
                            Check="Not Match"
                if Check=="Match":
                    print("\n"+("_").join(NProc)+" Check tool version is "+Fore.GREEN+"Match\n")
                else:
                    print("\n"+("_").join(NProc)+" Check tool version is "+Fore.RED+"Not Match\n")
    except ValueError:
        pass
        #print("\n"+"Table Format "+Fore.RED+"Error"+".\n")
    except:
        pass
        #print("\n"+"Check Toolversion "+Fore.RED+"Error"+".\n")

    #=================End=======================================================================
    print("End".center(90, "="))
    #======For Intel Check
    if Platform_Flag(OldProcPkgInfo) == "Intel" or Platform_Flag(OldProcPkgInfo) == "Intel G3":
        CheckPkg(NewProcPkgInfo, NewVersion)# Check new release Pkg is OK?
        PrintBiosBuildDate(Matchfolderlist, BiosBuildDate)
        PrintBiosBinaryCRC32(NewProcPkgInfo, BiosBinaryCRC32, NewVersion)
    #======For AMD Check
    else:
        CheckPkg_AMD(NewProcPkgInfo, NewVersion)# Check new release Pkg is OK?
        PrintBiosBuildDate(Matchfolderlist, BiosBuildDate)
        PrintBiosBinaryCRC32(NewProcPkgInfo, BiosBinaryCRC32, NewVersion)
    print("\nFinally pkg please compare with leading project.\n")
    os.system('Pause')
    sys.exit()
    # =================Script End===============================================================
