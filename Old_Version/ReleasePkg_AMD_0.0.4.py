# Make Release Pkg Script For AMD.   0.0.4
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com

# coding=UTF-8
from Lib import *
import sys, os
from colorama import init, Fore, Back, Style

# In add the Projects to be here, and the following BoradID should also add.
BoardID=["Q26", "Q27"]
ProjectName={"Q26":"ScottyRr", "Q27":"Scotty"}

Versionfilelist=["BUFF2.nsh"]

Removefilerule=["DCI.7z", ".7z", ".xlsm", ".cat", ".cer", ".pfx", ".pvk", \
                "oldversion"+".bin", "oldversion"+".inf", "oldversion"+".xml"]

ProductionReleaseServer={   "host":"ftp.usa.hp.com",
                            "username":"sign_ron",
                            "pwd":"5gg9*0UA"}

    # Script Start
if __name__ == '__main__':
    init(autoreset=True)
    multiprocessing.freeze_support()# For windows do multiprocessing.

    #=================Script Start==========================================================
    print("Make Release Pkg Script For AMD   0.0.4".center(90, "="))
    print("Input Information".center(90, "="))
    OldVersion=InputStr("OldVersion:")# Input OldVersion & NewVersion & NewBuildID
    NewVersion=InputStr("NewVersion:")
    NewBuildID=InputStr("NewBuildID:")# ex: 020106_"0001"
    ProcessProject=InputStr("                                   (Can Multiple Select)\
        \nPlease Enter Projects To Processed"+str(BoardID)+":")# Input need Process boardID, Can multiple choice
    ProcessProjectList=ProcessProject.upper().split()
    if OldVersion=="":print("\nPlease Input OldVersion.");sys.exit()
    elif NewVersion=="":print("\nPlease Select NewVersion.");sys.exit()

    #=================Find Need Process Old Pkg==============================================
    print("Find Need Process Old Pkg".center(90, "="))
    OldBuildID="0";NeedProcOldPkg=[]
    OldVersionFind = OldVersion[0:2]+"."+OldVersion[2:4]+"."+OldVersion[4:6]
    print("Start Find Process Project Old Pkg.")
    for Project in ProcessProjectList:
        temp=[]
        for Dir in os.listdir(".\\"):# Find oldveriosn pkg.
            if not Dir.split("_")[0]=="Fv" and not Dir.find(".7z")!=-1 and not Dir.find(".zip")!=-1:
                if Project+"_"+OldVersionFind in Dir:
                    temp.append(Dir)
        if len(temp)>1 and OldBuildID=="0":# If find oldveriosn pkg have buildID.
            OldBuildID=InputStr("\n"+str(temp)+"\n"+Project+"_"+OldVersionFind+" Please Select OldBuildID:")
            if OldBuildID=="" or OldBuildID=="0000":
                temp=[a for a in temp if len(a.split("_"))==2]
            else:
                temp=[a for a in temp if Project+"_"+OldVersionFind+"_"+OldBuildID in a]
            if not len(temp)==1:
                temp=[Dir for Dir in os.listdir(".\\") if Project+"_"+OldVersionFind in Dir]
                OldBuildID=InputStr("\n"+str(temp)+"\n"+Project+"_"+OldVersionFind+" Please Select OldBuildID:")
                if OldBuildID=="" or OldBuildID=="0000":
                    temp=[a for a in temp if len(a.split("_"))==2]
                else:
                    temp=[a for a in temp if Project+"_"+OldVersionFind+"_"+OldBuildID in a]
        if len(temp)==0:# Can't find oldveriosn pkg.
            print(Project+"_"+OldVersionFind+" Old Pkg folder can't find, Please check.")
        else:# Add find oldveriosn pkg.
            NeedProcOldPkg.append(temp[0])
    print("\nYour need process old Pkg:\n"+str(NeedProcOldPkg))

    #=================Make need process old/new Pkg info table===============================
    OldProcPkgInfo=[]
    OldProcPkgInfo=[Proc.split("_") for Proc in NeedProcOldPkg]# Split to process
    NewProcPkgInfo=[Proc.split("_") for Proc in NeedProcOldPkg]# New Pkg Name List
    for Proc in range(len(NeedProcOldPkg)):
        NewVersionFind = NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]
        NewProcPkgInfo[Proc][1] = NewVersionFind
        if not NewBuildID=="" or NewBuildID=="0000":
            NewProcPkgInfo[Proc].append(NewBuildID)

    #=================Find Fv Folder Or Zip File=============================================
    print("Find Fv Folder Or Zip File".center(90, "="))
    Matchfolderlist=FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)
    Matchziplist=FindFvZip(ProcessProjectList, NewVersion, NewBuildID)

    for Project in range(len(NewProcPkgInfo)):# ['Q26', '01.04.01']=>['ScottyRr', 'ScottyRr', 'Q26', '010401']
        NewProcPkgInfo[Project][1] = NewVersion
        NewProcPkgInfo[Project].insert(0, ProjectName[NewProcPkgInfo[Project][0]])
        NewProcPkgInfo[Project].insert(0, ProjectName[NewProcPkgInfo[Project][1]])

    print("Your Fv Folder: ", end=" ");print(Matchfolderlist)
    print("Your Fv Zip File: ", end=" ");print(str(Matchziplist)+"\n")
    # If can't find Fv folder or Zip file.
    if len(Matchfolderlist)==0 and len(Matchziplist)==0:
        print("Can't find Fv folder and zip file.\nDownload Fv files from Production Release FTP.\n")
        temp=Ftp_multi(NewProcPkgInfo, ProductionReleaseServer)[:]
        for name in temp:
            if str(name).find(".zip")!=-1:
                Matchziplist.append(name)
    # Number of Fv folders not match Projectlist
    elif len(Matchfolderlist)<len(ProcessProjectList) and len(Matchfolderlist)!=0 and len(Matchziplist)<len(ProcessProjectList):
        print("Number of Fv folders not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp=Ftp_multi(NewProcPkgInfo, ProductionReleaseServer)[:]
        for name in temp:
            if str(name).find(".zip")!=-1:
                Matchziplist.append(name)
    # Can't find Fv folders and Number of Fv Zip files not match Projectlist
    elif len(Matchfolderlist)==0 and len(Matchziplist)<len(ProcessProjectList):
        print("Fv Zip files not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
        temp=Ftp_multi(NewProcPkgInfo, ProductionReleaseServer)[:]
        for name in temp:
            if str(name).find(".zip")!=-1:
                Matchziplist.append(name)
    # If Fv folder already exists
    elif len(Matchfolderlist)==len(ProcessProjectList):
        print("Fv folder already exists.\n")
    # Find Fv Zip file Start extracting

    if len(Matchziplist)!=-1:
        print("\nFind Fv Zip File, Start Extracting.")
        for Zipname in Matchziplist:
            if not os.path.isdir(".\\"+Zipname.split(".")[0]):
                UnZip(Zipname)
                shutil.move(".\\"+Zipname, ".\\"+Zipname.split(".")[0])
                print(Zipname+" Extract succeeded.")
                Matchfolderlist.append(Zipname.split(".")[0])
            else:
                print("Fv folder "+Zipname.split(".")[0]+" already exists, Remove Fv zip file.")
                os.remove(".\\"+Zipname)
        print("\nNow Your Fv Folder: ", end=" ");print(Matchfolderlist)
    else:   print("\nNow Your Fv Folder: ", end=" ");print(Matchfolderlist)

    #=================Find New Pkg Or Add New Pkg=============================================
    print("Find New Pkg Or Add New Pkg".center(90, "="))

    for project in range(len(NewProcPkgInfo)):# ['ScottyRr', 'ScottyRr', 'Q26', '010401']=>['Q26', '01.04.01']
        NewProcPkgInfo[project].remove(ProjectName[NewProcPkgInfo[project][2]])
        NewProcPkgInfo[project].remove(ProjectName[NewProcPkgInfo[project][1]])
        NewProcPkgInfo[project][1] = NewVersion[0:2]+"."+NewVersion[2:4]+"."+NewVersion[4:6]

    for OProc in range(len(OldProcPkgInfo)):# How much Old Version folder
        OldVersionPath=".\\"+("_").join(OldProcPkgInfo[OProc])
        NewVersionPath=".\\"+("_").join(NewProcPkgInfo[OProc])
        if not os.path.isdir(NewVersionPath):# Check NewVersion Folder Exist
            if not os.path.isdir(OldVersionPath+"\\"+("_").join(OldProcPkgInfo[OProc])):# Check Old Pkg is in folder???
                Copytree(OldVersionPath, NewVersionPath)
            elif not os.path.isdir(OldVersionPath+"\\FPTW"):
                if os.path.isdir(OldVersionPath+"\\"+("_").join(OldProcPkgInfo[OProc])):# Check Old Pkg is in folder???
                    Copytree(OldVersionPath+"\\"+("_").join(OldProcPkgInfo[OProc]), NewVersionPath)
                else:   print("Pkg "+("_").join(OldProcPkgInfo[OProc])+" can't find.")
            else:   print("Pkg "+("_").join(OldProcPkgInfo[OProc])+" can't find.")
        else:   print("Pkg "+NewVersionPath.split("\\")[-1]+" already exists.")
    if len(Matchfolderlist)==0: print("Can't find anything Fv folder.\n")

    #=================Modify Pkg Update Version==============================================
    print("Modify Pkg Update Version".center(90, "="))
    for NProc in NewProcPkgInfo:# Pkg Modify Update Version
        Path=os.getcwd()+"\\"+("_").join(NProc)
        if os.path.isdir(Path+"\\AMDFLASH"):# Check Folder Exist
            ReleaseNote=[ReleaseNote for ReleaseNote in os.listdir(Path) if "Release_Notes.docx" in ReleaseNote]
            if len(ReleaseNote)==1:
                os.chdir(Path)
                os.rename(ReleaseNote[0], "Scotty_"+("_").join(NProc)+"_Release_Notes.docx")
                if os.path.isfile("Scotty_"+("_").join(NProc)+"_Release_Notes.docx"):
                    print("ReleaseNote Rename to "+"Scotty_"+("_").join(NProc)+"_Release_Notes.docx"+" succeeded.")
            else:
                print("Can't find Release_Notes.docx")
            os.chdir(Path+"\\AMDFLASH")
            for Filename in Versionfilelist:
                File=open(Filename,"r+")
                Strlist=File.read().split("/")
                if "all "+NProc[0]+"_"+OldVersion+".bin " in Strlist:
                    for Str in range(len(Strlist)):
                        if Strlist[Str]=="all "+NProc[0]+"_"+OldVersion+".bin ":
                            Strlist[Str]="all "+NProc[0]+"_"+NewVersion+".bin "
                elif NewVersion in Strlist:
                    print("Now Pkg Version is already "+Strlist[1]+".")
                    File.close()
                    break
                else:
                    print("Old Version "+OldVersion+" Can't find, Now Pkg Version is "+Strlist[1].split("_")[1]+".")
                    File.close()
                    break
                Strlist="/".join(Strlist)
                File.seek(0,0)
                File.write(Strlist)
                File.close()
            print("Modify "+str(NProc[0])+" Pkg Version (.nsh & .bat) to "+str(NProc[1])+" succeeded.\n")
        else:   print("Pkg "+("_").join(NProc)+" can't find.\n")
        os.chdir("..\..")

    #=================Remove Pkg Old File=====================================================
    print("Remove Pkg Old File".center(90, "="))
    for NProc in NewProcPkgInfo:
        targetfolder=("_").join(NProc)
        if (os.path.isdir(".\\"+targetfolder)):
            RemoveOldFileInDir(targetfolder, OldVersion, Removefilerule)
        else:   print("Pkg "+("_").join(NProc)+" can't find.\n")

    #=================Fv File Rename And Copy To Pkg===========================================
    print("Fv File Rename And Copy To Pkg".center(90, "="))
    for Fv in Matchfolderlist:
        for NProc in NewProcPkgInfo:
            if Fv.split("_")[1]==NProc[0]:
                if os.path.isdir(".\\"+Fv):
                    Path=os.getcwd()+"\\"+Fv
                    Boardversion=NProc[0]+"_"+NewVersion
                    if os.path.isfile(Path+"\\"+Boardversion+".bin") or os.path.isfile(Path+"\\"+Boardversion+"_16.bin"):# For 16MB BIOS
                        if os.path.isfile(Path+"\\"+Boardversion+".inf") and os.path.isfile(Path+"\\"+NProc[0]+".xml"):
                            if os.path.isdir(".\\"+("_").join(NProc)):# Check Pkg Folder Exist
                                CopyFiles_AMD(Fv, ("_").join(NProc), NewVersion)
                            else:   print("Pkg "+("_").join(NProc)+" can't find.")
                else:   print("Need to be processed Fv folder:"+Fv+" can't find.\n")
    if len(Matchfolderlist)==0 or len(ProcessProjectList)==0:  print("Can't find anything Fv folder.\n")

    #=================End=======================================================================
    print("End".center(90, "="))
    CheckPkg_AMD(NewProcPkgInfo, NewVersion)# Check new release Pkg is OK?
    print("\nFinally pkg please compare with leading project.\n")
    os.system('Pause')
    sys.exit()
    # Script End
