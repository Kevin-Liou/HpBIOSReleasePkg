# Make Release Pkg Script.   3.0.0
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com

# coding=UTF-8
from Lib import *
import sys, os
from colorama import init, Fore, Back, Style

# In add the Projects to be here, and the following BoradID should also add.
BoardID=["Q10", "Q11", "Q21", "Q22", "Q23", "R10", "R11", "R21", "R22", "R23"]

Versionfilelist=["BUFF2.nsh", "Buff2All.nsh", "Update32.bat", "Update64.bat", "UpdateEFI.nsh"]

Removefilerule=["DCI.7z", ".7z", ".xlsm", ".cat", ".cer", ".pfx", ".pvk", "_12.bin", "_16.bin",\
                "_32.bin", "oldversion"+".bin", "oldversion"+".inf", "oldversion"+".xml"]

ProductionReleaseServer={   "host":"ftp.usa.hp.com",
                            "username":"sign_ron",
                            "pwd":"5gg9*0UA"}

    # Script Start
if __name__ == '__main__':
    init(autoreset=True)
    multiprocessing.freeze_support()# For windows do multiprocessing.

    #=================Script Start==========================================================
    print("Make Release Pkg Script   3.0.0".center(90, "="))
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
    print("Start Find Process Project Old Pkg.")
    for Project in ProcessProjectList:
        temp=[]
        for Dir in os.listdir(".\\"):# Find oldveriosn pkg.
            if not Dir.split("_")[0]=="Fv" and not Dir.find(".7z")!=-1 and not Dir.find(".zip")!=-1:
                if Project+"_"+OldVersion in Dir:
                    temp.append(Dir)
        if len(temp)>1 and OldBuildID=="0":# If find oldveriosn pkg have buildID.
            OldBuildID=InputStr("\n"+str(temp)+"\n"+Project+"_"+OldVersion+" Please Select OldBuildID:")
            if OldBuildID=="" or OldBuildID=="0000":
                temp=[a for a in temp if len(a.split("_"))==4]
            else:
                temp=[a for a in temp if Project+"_"+OldVersion+"_"+OldBuildID in a]
            if not len(temp)==1:
                temp=[Dir for Dir in os.listdir(".\\") if Project+"_"+OldVersion in Dir]
                OldBuildID=InputStr("\n"+str(temp)+"\n"+Project+"_"+OldVersion+" Please Select OldBuildID:")
                if OldBuildID=="" or OldBuildID=="0000":
                    temp=[a for a in temp if len(a.split("_"))==4]
                else:
                    temp=[a for a in temp if Project+"_"+OldVersion+"_"+OldBuildID in a]
        if len(temp)==0:# Can't find oldveriosn pkg.
            print(Project+"_"+OldVersion+" Old Pkg folder can't find, Please check.")
        else:# Add find oldveriosn pkg.
            NeedProcOldPkg.append(temp[0])
    print("\nYour need process old Pkg:\n"+str(NeedProcOldPkg))

    #=================Make need process old/new Pkg info table===============================
    OldProcPkgInfo=[]
    OldProcPkgInfo=[Proc.split("_") for Proc in NeedProcOldPkg]# Split to process
    NewProcPkgInfo=OldProcPkgInfo[:]# New Pkg Name List
    for Proc in range(len(NewProcPkgInfo)):
        NewProcPkgInfo[Proc]=NewProcPkgInfo[Proc][:3]
        NewProcPkgInfo[Proc].append(NewVersion)
        if not NewBuildID=="" or NewBuildID=="0000":
            NewProcPkgInfo[Proc].append(NewBuildID)

    #=================Find Fv Folder Or Zip File=============================================
    print("Find Fv Folder Or Zip File".center(90, "="))
    Matchfolderlist=FindFvFolder(ProcessProjectList, NewVersion, NewBuildID)
    Matchziplist=FindFvZip(ProcessProjectList, NewVersion, NewBuildID)

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
        Path=".\\"+("_").join(NProc)
        if os.path.isdir(Path+"\\FPTW"):# Check Folder Exist
            ReleaseNote=[ReleaseNote for ReleaseNote in os.listdir(Path) if "release note.docx" in ReleaseNote]
            if len(ReleaseNote)==1:
                os.chdir(Path)
                os.rename(ReleaseNote[0], ("_").join(NProc)+" release note.docx")
            print("ReleaseNote Rename to "+("_").join(NProc)+" release note.docx"+" succeeded.")
            os.chdir(".\\FPTW")
            ChangeBuildID(NProc, Versionfilelist, OldVersion, NewVersion)
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
            if Fv.split("_")[1]==NProc[2]:
                if os.path.isdir(".\\"+Fv):
                    Path=".\\"+Fv
                    Boardversion=NProc[2]+"_"+NProc[3]
                    if os.path.isfile(".\\"+Fv+"\\"+NProc[2]+"_"+NProc[3]+"_16.bin"):# For 16MB BIOS
                        if os.path.isfile(Path+"\\"+Boardversion+"_12.bin") and os.path.isfile(Path+"\\"+Boardversion+".xml"):# If Alreadly renamed
                            print(Boardversion+"_12.bin & "+Boardversion+".xml alreadly renamed.")
                        if os.path.isfile(Path+"\\"+Boardversion+".bin") and os.path.isfile(Path+"\\"+NProc[2]+".xml"):
                            os.rename(Path+"\\"+Boardversion+".bin", Path+"\\"+Boardversion+"_12.bin")# Rename Fv folder 2 files
                            os.rename(Path+"\\"+NProc[2]+".xml", Path+"\\"+Boardversion+".xml")
                            print(Boardversion+"_12.bin & "+Boardversion+".xml rename succeeded.")
                        if os.path.isfile(Path+"\\"+Boardversion+".inf"):
                            if os.path.isdir(".\\"+("_").join(NProc)):# Check Pkg Folder Exist
                                CopyFiles(Fv, ("_").join(NProc))
                            else:   print("Pkg "+("_").join(NProc)+" can't find.")
                    else:
                        if os.path.isfile(Path+"\\"+Boardversion+"_16.bin") and os.path.isfile(Path+"\\"+Boardversion+".xml"):# If Alreadly renamed
                            print(Boardversion+"_16.bin & "+Boardversion+".xml alreadly renamed.")
                        if os.path.isfile(Path+"\\"+Boardversion+".bin") and os.path.isfile(Path+"\\"+NProc[2]+".xml"):
                            os.rename(Path+"\\"+Boardversion+".bin", Path+"\\"+Boardversion+"_16.bin")# Rename Fv folder 2 files
                            os.rename(Path+"\\"+NProc[2]+".xml", Path+"\\"+Boardversion+".xml")
                            print(Boardversion+"_16.bin & "+Boardversion+".xml rename succeeded.")
                        if os.path.isfile(Path+"\\"+Boardversion+".inf"):
                            if os.path.isdir(".\\"+("_").join(NProc)):# Check Pkg Folder Exist
                                CopyFiles(Fv, ("_").join(NProc))
                            else:   print("Pkg "+("_").join(NProc)+" can't find.")
                else:   print("Need to be processed Fv folder:"+Fv+" can't find.\n")
    if len(Matchfolderlist)==0 or len(ProcessProjectList)==0:  print("Can't find anything Fv folder.\n")

    #=================Check Tool Version Is Match In Table======================================
    print("Check Tool Version Is Match In Table".center(90, "=")+"\n")
    try:
        for NProc in NewProcPkgInfo:
            Toolversiontablepath=".\\"+("_").join(NProc)+"\\FactoryUtility\\ToolVersion.xlsx"
            Toolversioninfo=ReadToolVersionTable(Toolversiontablepath)
            Check="Match"
            print("Name".ljust(30)+"File".ljust(14)+"Table".ljust(20)+"File".ljust(13)+"Table")
            for name, type, verinfo, path, none, dateinfo in Toolversioninfo:
                ver=ChangeVersionInfo(verinfo)
                date=ChangeDataInfo(dateinfo)
                fileinfo_temp=GetFileInfo(".\\"+("_").join(NProc)+path+name)
                if name=="EEUPDATEW32.exe" or name=="EEUPDATEW64e.exe":
                    if (os.listdir(".\\"+("_").join(NProc)+"\\Global\\BIOS")[0].split("_")[2]=="16.bin"):# If 400 project
                        continue
                    if fileinfo_temp[0]!=ver or fileinfo_temp[1]!=date:
                        print(name.ljust(30)+fileinfo_temp[0]+"==>"+str(ver)+"\t\t"+fileinfo_temp[1]+"==>"+str(date))
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
                        print(name.ljust(30)+fileinfo_temp[0]+"==>"+ver+"\t\t"+fileinfo_temp[1]+"==>"+date)
                        Check="Not Match"
            if Check=="Match":
                print("\n"+("_").join(NProc)+"Check tool version is "+Fore.GREEN+"Match\n")
            else:
                print("\n"+("_").join(NProc)+"Check tool version is "+Fore.RED+"Not Match\n")
    except ValueError:
        print("\n"+"Table Format "+Fore.RED+"Error"+".\n")
    except:
        print("\n"+"Check Toolversion "+Fore.RED+"Error"+".\n")

    #=================End=======================================================================
    print("End".center(90, "="))
    CheckPkg(NewProcPkgInfo, NewVersion)# Check new release Pkg is OK?
    print("\nFinally pkg please compare with leading project.\n")
    os.system('Pause')

    sys.exit()
    # Script End