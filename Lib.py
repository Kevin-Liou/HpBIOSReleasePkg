#=================================
# coding=UTF-8
# Make Release Pkg Script Library.   1.8.13
# Author: Kevin Liou
# Contact: Kevin.Liou@quantatw.com

#This script is using "Make Release Pkg Script"   4.7.21
#=================================

from multiprocessing import Pool, Manager, freeze_support
from colorama import init, Fore
from shutil import copy, copytree, move, rmtree
from win32com.client import Dispatch
from ftplib import FTP, error_perm
from datetime import datetime
from zipfile import ZipFile
from re import sub, search
from time import sleep, localtime, strftime
from xlwings import constants

import sys, os, glob, logging, argparse
import xlwings as xw
import struct


def argparse_function(ver):
    parser = argparse.ArgumentParser(prog='ReleasePkg.py', description='Tutorial')
    parser.add_argument("-d", "--debug", help='Show debug message.', action="store_true")
    parser.add_argument("-v", "--version", action="version", version=ver)
    args = parser.parse_args()
    if args.debug:
        Debug_Format = "%(levelname)s, %(funcName)s: %(message)s"
        logging.basicConfig(level=logging.DEBUG, format=Debug_Format)  #Debug use flag
        print("Enable debug mode.")
    return ver

# Input Str, "Str" is message.
def InputStr(Str):
    if sys.version > '3':
        OutStr = input(Str)
    else:
        OutStr = raw_input(Str)
    return OutStr


# Check platform is Intel or AMD.
def Platform_Flag(IDCheck):
    if type(IDCheck) == list:
        for ID in IDCheck:
            if "Q26" in ID:
                return "Q26"
            if "Q27" in ID:
                return "Q27"
            if "R26" in ID:
                return "R26"
            if "R24" in ID:
                return "R24"
            if "S25" in ID:
                return "S25"
            if "S27" in ID:
                return "S27"
            if "S29" in ID:
                return "S29"
            if "T25" in ID:
                return "T25"
            if "T26" in ID:
                return "T26"
            if "T27" in ID:
                return "T27"
            if "P10" in ID:
                return "Intel G3"
            if ("Q10" in ID) or ("Q11" in ID) or ("Q21" in ID) or ("Q22" in ID) or ("Q23" in ID) or ("Q35" in ID):
                return "Intel G4"
            if ("R11" in ID) or ("R21" in ID) or ("R22" in ID) or ("R23" in ID):
                return "Intel G5"
            if ("S10" in ID) or ("S11" in ID) or ("S21" in ID) or ("S22" in ID) or ("S23" in ID):
                return "Intel G6"
            if ("T11" in ID) or ("T21" in ID) or ("T22" in ID):
                return "Intel G8"
            if ("U11" in ID) or ("U21" in ID) or ("U22" in ID) or ("U23" in ID):
                return "Intel G9"
            if ("V11" in ID) or ("V21" in ID) or ("V22" in ID) or ("V23" in ID):
                return "Intel G10"
    else:
        if "Q26" in IDCheck:
            return "Q26"
        if "Q27" in IDCheck:
            return "Q27"
        if "R26" in IDCheck:
            return "R26"
        if "R24" in IDCheck:
            return "R24"
        if "S25" in IDCheck:
            return "S25"
        if "S27" in IDCheck:
            return "S27"
        if "S29" in IDCheck:
            return "S29"
        if "T25" in IDCheck:
            return "T25"
        if "T26" in IDCheck:
            return "T26"
        if "T27" in IDCheck:
            return "T27"
        if "P10" in IDCheck:
            return "Intel G3"
        if ("Q10" in IDCheck) or ("Q11" in IDCheck) or ("Q21" in IDCheck) or ("Q22" in IDCheck) or ("Q23" in IDCheck) or ("Q35" in IDCheck):
            return "Intel G4"
        if ("R11" in IDCheck) or ("R21" in IDCheck) or ("R22" in IDCheck) or ("R23" in IDCheck):
            return "Intel G5"
        if ("S10" in IDCheck) or ("S11" in IDCheck) or ("S21" in IDCheck) or ("S22" in IDCheck) or ("S23" in IDCheck):
            return "Intel G6"
        if ("T11" in IDCheck) or ("T21" in IDCheck) or ("T22" in IDCheck):
            return "Intel G8"
        if ("U11" in IDCheck) or ("U21" in IDCheck) or ("U22" in IDCheck) or ("U23" in IDCheck):
            return "Intel G9"
        if ("V11" in IDCheck) or ("V21" in IDCheck) or ("V22" in IDCheck) or ("V23" in IDCheck):
            return "Intel G10"


# path:Unzip file path.
def UnZip(path):
    with ZipFile(path, 'r') as myzip:
        for file in myzip.namelist():
            myzip.extract(file,path.rstrip(".zip"))
        myzip.close()


def PrintZipAllInfo(path):
    zf = ZipFile(path)
    for info in zf.infolist():
        print(info.filename)
        print('\tComment:\t' + str(info.comment))
        print('\tModified:\t' + str(datetime(*info.date_time)))
        print('\tZIP version:\t' + str(info.create_version))
        print('\tCompressed:\t' + str(info.compress_size) + 'bytes')
        print('\tUncompressed:\t' + str(info.file_size) + 'bytes')


def GetZipDateInfo(path, name):
    zf = ZipFile(path)
    for info in zf.infolist():
        if name in info.filename:
            return datetime(*info.date_time).strftime("%Y/%m/%d")


# Working with multiple folders.
def MatchMultipleFolder(Matchfolderlist):
    print("Check is multiple folder?")
    for Fv in Matchfolderlist:
        for root, dirs, files in os.walk(".\\" + Fv):
            for file in files:
                if file == "AutoGenFlashMap.h":
                    if (root + "\\" + file) == (".\\" + Fv + "\\" + "AutoGenFlashMap.h"):
                        print("No, not need move.")
                        return
                    else:
                        print("Yes, need move.")
                        for folder in dirs:
                            move(root + "\\" + folder, ".\\" + Fv + "\\")
                        for file in files:
                            move(root + "\\" + file, ".\\" + Fv + "\\")
                        if os.path.exists(".\\" + Fv + "\\" + "AutoGenFlashMap.h"):
                            return
                        else:
                            print("Can't find AutoGenFlashMap.h in Fv folder, Please check Fv folder.")
                            sys.exit()


# Modify Update Version message in file.
def ChangeBuildID(NewProcPkgInfo, Versionfilelist, NewVersion):
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
    for Filename in Versionfilelist:
        if os.path.isfile(Filename):
            File = open(Filename,"r+")
            Strlist = File.read()
            ReStrlist = sub(pattern, PlatID + "_" + NewVersion, Strlist)
            if ReStrlist != Strlist:
                File.seek(0, 0)
                File.write(ReStrlist)
                File.close()
            else:
                print("Now Pkg Version is already " + str(NewVersion) + ".")
                File.close()
                break
    print("Modify " + ("_").join(NewProcPkgInfo) + " Pkg Version (.nsh & .bat) to " + str(NewVersion) + " succeeded.\n")


# Remove old file in pkg.
def RemoveOldFileInDir(targetdir, RemoveRule, NotRemoveRule):
    i = 0
    for root,dirs,files in os.walk(targetdir):
        for name in files:# Here are the rules for remove
            Path = ".\\" + os.path.join(root, name)
            for Rule_remove in RemoveRule:
                pattern_remove = Rule_remove
                searchObj_remove = search(pattern_remove, Path)
                if os.path.isfile(Path) and searchObj_remove != None: # If Rule_remove in name
                    for Rule_notremove in NotRemoveRule:
                        pattern_notremove = Rule_notremove
                        searchObj_notremove = search(pattern_notremove, Path)
                        if searchObj_notremove != None: # If searchObj_notremove != None , Break, Stop comparing and remove
                            break
                    if searchObj_notremove == None:
                        os.remove(Path)
                        i = i + 1
                        print(os.path.join(root, name) + "\t remove succeeded.")
    if i == 0:
        print(targetdir+"\t no file can remove.")
    print()


# Copy old version folder to new version folder.
def Copy_Release_Folder(sourcePath, targetPath):
    print("Start Copy "+sourcePath.split("\\")[-1] + " to "+targetPath.split("\\")[-1] + ", Please wait.....")
    copytree(sourcePath, targetPath)# Copy to new Pkg
    print("Copy Pkg " + sourcePath.split("\\")[-1] + " to " + targetPath.split("\\")[-1] + " succeeded.\n")


# Not use now.
# If Fv folder is new folder.
def New_FvFolder_Move_File(Fv_Path):
    if os.path.isdir(Fv_Path + "\\Combined\\WU") and os.path.isfile(Fv_Path + "\\Combined\\WU\\fwu.pvk"):
        for root, dirs, files in os.walk(Fv_Path + "\\Combined\\WU"):
            for name in files:# Move "WU" file.
                if os.path.isfile(Fv_Path + "\\Combined\\WU\\" + name):
                    print("move " + Fv_Path + "\\Combined\\WU\\" + name + " to " + Fv_Path + "\\Combined")
                    move(Fv_Path + "\\Combined\\WU\\" + name, Fv_Path + "\\Combined")
                elif os.path.isfile(Fv_Path + "Combined\\WU\\" + name):
                    print("move " + Fv_Path + "Combined\\WU\\" + name + " to " + Fv_Path + "\\Combined")
                    move(Fv_Path + "Combined\\WU\\" + name, Fv_Path + "\\Combined")


# Cpoy Fv folder file to NewPkg.
def Copy_Release_Files(sourceFolder, targetFolder, NProc, Matchfolderlist):
    sourcefullpath = ".\\" + sourceFolder + "\\"
    targetfullpath = ".\\" + targetFolder + "\\"
    # Combined copy to Capsule&HPFWUPDREC.
    #======For G5 and late Fv
    # Check Capsule folder format for G5 and late.
    if (Platform_Flag(targetFolder) == "Intel G5") or (Platform_Flag(targetFolder) == "Intel G6") or\
        (Platform_Flag(targetFolder) == "Intel G8") or (Platform_Flag(targetFolder) == "Intel G9") or\
        (Platform_Flag(targetFolder) == "Intel G10"):
        if os.path.isdir(targetfullpath + "\\Capsule\\CCG5") and ((Platform_Flag(targetFolder) == "Intel G6") or \
            (Platform_Flag(targetFolder) == "Intel G8") or (Platform_Flag(targetFolder) == "Intel G9") or \
            (Platform_Flag(targetFolder) == "Intel G10")):
            os.rename(targetfullpath + "\\Capsule\\CCG5", targetfullpath + "\\Capsule\\PD_FW")
        if not os.path.isdir(targetfullpath + "\\Capsule\\Windows"):
            os.makedirs(targetfullpath + "\\Capsule\\Windows")
            os.makedirs(targetfullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
            os.makedirs(targetfullpath + "\\Capsule\\Windows\\Thunderbolt")
        if not os.path.isdir(targetfullpath + "\\Capsule\\Linux"):
            os.makedirs(targetfullpath + "\\Capsule\\Linux")
            os.makedirs(targetfullpath + "\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)")
        for file in glob.glob(targetfullpath + "\\Capsule\*.doc*"):
            if file.find("submission") != -1 or file.find("Submission") != -1:
                move(file, targetfullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
                print("move " + file + " to " + targetfullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
        # Copy FUR and WU files.
        if os.path.isdir(sourcefullpath + "\\Combined\\FUR") and os.path.isdir(sourcefullpath + "\\Combined\\WU"):
            for root,dirs,files in os.walk(sourcefullpath + "\\Combined\\FUR"):
                for name in files:
                    if name.find(".bin") != -1 or name.find(".inf") != -1:
                        copy(root + "\\" + name, targetfullpath + "\\HPFWUPDREC")
                        print(root + "\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
            for root,dirs,files in os.walk(sourcefullpath + "\\Combined\\WU"):
                for name in files:
                    copy(root + "\\" + name, targetfullpath + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)" + " Copy succeeded.")
        # If Linux folder exist, copy files.
        if os.path.isdir(sourcefullpath+"\\Combined\\Linux"):
            for root,dirs,files in os.walk(sourcefullpath+"\\Combined\\Linux"):
                for name in files:
                    copy(root + "\\" + name, targetfullpath + "\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)" + " Copy succeeded.")
        # If TBT FW exist, copy it, for Intel G5 and late if support TBT use.
        TBT_path = ""
        TBT_Version = ""
        if not os.path.exists(targetfullpath + "\\Capsule\\Windows\\Thunderbolt"):
            os.makedirs(targetfullpath + "\\Capsule\\Windows\\Thunderbolt")
        if os.path.isdir(sourcefullpath + "\\TBT"):
            TBT_path = sourcefullpath + "\\TBT"
        elif os.path.isdir(sourcefullpath + "\\TBT_RTD3"):
            TBT_path = sourcefullpath + "\\TBT_RTD3"
        if not TBT_path == "":
            for root,dirs,files in os.walk(TBT_path):
                for name in files:
                    if os.path.isfile(targetfullpath + "\\Capsule\\Windows\\Thunderbolt\\" + name):
                        os.remove(targetfullpath + "\\Capsule\\Windows\\Thunderbolt\\" + name)
                    copy(root + "\\" + name, targetfullpath + "\\Capsule\\Windows\\Thunderbolt")
                    print(root + "\\" + name + " to " + targetfullpath + "\\Capsule\\Windows\\Thunderbolt" + " Copy succeeded.")
                    if ".inf" in name:
                        Pattern = r"HP_\w+_\w+_\w+_\w+_\w+_\w+_\d+"
                        File = open(root + "\\" + name,"r+")
                        Strlist = File.read()
                        searchObj = search(Pattern, Strlist)
                        if searchObj != None:
                            TBT_Version = searchObj.group(0)
                            File.close()
            if TBT_Version != "":
                Dirs = os.listdir(targetfullpath + "\\Capsule\\Windows\\Thunderbolt")
                for file in Dirs:
                    if ".bin" in file:
                        os.rename(targetfullpath + "\\Capsule\\Windows\\Thunderbolt\\" + file, targetfullpath + "\\Capsule\\Windows\\Thunderbolt\\" + TBT_Version + ".bin")
                    elif ".inf" in file:
                        os.rename(targetfullpath + "\\Capsule\\Windows\\Thunderbolt\\" + file, targetfullpath + "\\Capsule\\Windows\\Thunderbolt\\" + TBT_Version + ".inf")
            if os.path.exists(targetfullpath + "\\Capsule\\TBT"):
                rmtree(targetfullpath + "\\Capsule\\TBT")
        # ME binary copy to METools\FWUpdate\HPSignME for Intel G5 and late.
        MEbinary_pattern = r"ME_+[0-9]+[\.]+[0-9]+[\.]+[0-9]+[\.]+[0-9]+.bin"
        if not os.path.exists(targetfullpath + "\\METools\\FWUpdate\\HPSignME"): # Copy sign ME file
            os.makedirs(targetfullpath + "\\METools\\FWUpdate\\HPSignME")
            logging.debug('Make dirs \\METools\\FWUpdate\\HPSignME.')
        ME_Bin_Check = "False"
        for root,dirs,files in os.walk(sourcefullpath + "\\ME"):
            for name in files:
                searchObj = search(MEbinary_pattern, name)
                if (searchObj != None):
                    copy(root + "\\" + name, targetfullpath + "\\METools\\FWUpdate\\HPSignME")
                    print(root + "\\" + name + "(Sign) to " + targetFolder + "\\METools\\FWUpdate\\HPSignME" + " Copy succeeded.")
                    ME_Version = CheckMEVersion(NProc, Matchfolderlist) # ex. 14.0.21.7227
                    ME_Bin_Check = "True"
                    logging.debug('ME_Version:' + ME_Version)
                    break
        if (ME_Bin_Check == "False") and (os.path.isfile(sourcefullpath + "\\ME\\ME.bin")):
            ME_filename = "ME.bin"
            if os.path.isfile(sourcefullpath + "\\ME\\ME.inf") and os.path.isfile(sourcefullpath + "\\ME\\ME.bin"):
                ME_Version = CheckMEVersion(NProc, Matchfolderlist) # ex. 14.0.21.7227
                logging.debug('ME_Version:' + ME_Version)
                if not os.path.isfile(sourcefullpath + "\\ME\\ME_" + ME_Version + ".bin"):
                    os.rename(sourcefullpath + "\\ME\\ME.bin", sourcefullpath + "\\ME\\ME_" + ME_Version + ".bin")
                    print("Rename: ME.bin to ME_" + ME_Version + ".bin")
                    ME_filename = "ME_" + ME_Version + ".bin"
            copy(sourcefullpath + "\\ME\\" + ME_filename, targetfullpath + "\\METools\\FWUpdate\\HPSignME")
            print(root + "\\" + ME_filename + "(Sign) to " + targetFolder + "\\METools\\FWUpdate\\HPSignME" + " Copy succeeded.")
        if os.path.isfile(sourcefullpath + "\\ME\\ME_0101.bin"):# Copy unsign ME file
            copy(sourcefullpath + "\\ME\\ME_0101.bin", targetfullpath + "\\METools\\FWUpdate\\MEFW")
            os.rename(targetfullpath + "\\METools\\FWUpdate\\MEFW\\ME_0101.bin", targetfullpath + "\\METools\\FWUpdate\\MEFW\\ME_" + ME_Version + ".bin")
            print(sourceFolder + "\\ME\\" + ME_Version + "(UnSign) to " + targetFolder + "\\METools\\FWUpdate\\MEFW" + " Copy succeeded.")
    #=======For G4 other Fv
    if (Platform_Flag(targetFolder) == "Intel G4"):
        if os.path.isfile(sourcefullpath + "\\Combined\\fwu.pfx"):
            for root,dirs,files in os.walk(sourcefullpath + "\\Combined"):
                for name in files:
                    if name.find(".bin") != -1 or name.find(".inf") != -1:
                        copy(root + "\\" + name, targetfullpath + "\\HPFWUPDREC")
                        print(root + "\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
                    copy(root + "\\" + name, targetfullpath  + "\\Capsule")
                    print(root + "\\" + name + " to " + targetFolder + "\\Capsule" + " Copy succeeded.")
    #======For G3 FV
    if Platform_Flag(targetFolder) == "Intel G3":
        if os.path.isfile(sourcefullpath + "\\fwu.pfx"):
            for root,dirs,files in os.walk(sourcefullpath):
                for name in files:
                    if name.find("_12.bin") != -1:
                        copy(root + "\\" + name, targetfullpath + "\\HPBIOSUPDREC")
                        os.rename(targetfullpath + "\\HPBIOSUPDREC\\" + name, targetfullpath + "\\HPBIOSUPDREC\\" + name[0:8] + ".bin")
                        print(sourceFolder + "\\" + name[0:8] + ".bin" + " to " + targetFolder + "\\HPBIOSUPDREC" + " Copy succeeded.")
                        copy(root + "\\" + name, targetfullpath + "\\Capsule Update")
                        os.rename(targetfullpath + "\\Capsule Update\\" + name, targetfullpath + "\\Capsule Update\\" + name[0:8] + ".bin")
                        print(sourceFolder + "\\" + name[0:8] + ".bin" + " to " + targetFolder + "\\Capsule Update" + " Copy succeeded.")
                    if name.find(".inf") != -1:
                        copy(root + "\\" + name, targetfullpath + "\\HPBIOSUPDREC")
                        print(sourceFolder + "\\" + name + " to " + targetFolder + "\\HPBIOSUPDREC" + " Copy succeeded.")
                    if name.find(".cer") != -1 or name.find(".pfx") != -1 or name.find(".pvk") != -1 or name.find(".cat" ) != -1 or name.find(".inf") != -1:
                        copy(root + "\\" + name, targetfullpath + "\\Capsule Update")
                        print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Capsule Update" + " Copy succeeded.")
    # Bin file copy to FPTW&Global.
    for name in os.listdir(sourcefullpath):
        if name.find("_12.bin") != -1 or name.find("_16.bin") != -1 or name.find("_32.bin") != -1:
            copy(sourcefullpath + name, targetfullpath + "\\FPTW")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\FPTW" + " Copy succeeded.")
        if name.find("_12.bin") != -1:# If 16MB BIOS case please add it.
            name = name.split("_12.bin")[0] + "_16.bin"
            copy(sourcefullpath + name, targetfullpath + "\\Global\\BIOS")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.")
        elif name.find("_32.bin") != -1:
            copy(sourcefullpath + name, targetfullpath + "\\Global\\BIOS")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.")
        if name.find(".xml") != -1:# Copy to XML
            if name.find(str(sourcefullpath.split("_")[1])) != -1:
                copy(sourcefullpath + name, targetfullpath + "\\XML")
                print(sourceFolder + "\\" + name + " to "+targetFolder + "\\XML" + " Copy succeeded.")
        # For Smart flash copy *Pvt.bin.
        if name.find("Pvt.bin") != -1:
            copy(sourcefullpath + name, targetfullpath)
            print(sourceFolder + "\\" + name + " to " + targetFolder + " Copy succeeded.")
    print()


# Cpoy Fv folder file to NewPkg.(For AMD)
def Copy_Release_Files_AMD(sourceFolder, targetFolder, NewVersion):
    sourcefullpath = ".\\" + sourceFolder + "\\"
    targetfullpath = ".\\" + targetFolder + "\\"
    # Combined copy to Capsule&HPFWUPDREC.
    #======For G5 and late Fv
    if os.path.isdir(sourcefullpath + "\\Combined\\FUR") and os.path.isdir(sourcefullpath + "\\Combined\\WU"):
        for root,dirs,files in os.walk(sourcefullpath + "\\Combined\\FUR"):
            for name in files:
                if name.find(".bin") != -1 or name.find(".inf") != -1:
                    copy(sourcefullpath + "Combined\\FUR\\" + name, targetfullpath + "\\HPFWUPDREC")
                    print(sourceFolder + "\\Combined\\FUR\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
        for root,dirs,files in os.walk(sourcefullpath + "\\Combined\\WU"):
            for name in files:
                copy(sourcefullpath + "\\Combined\\WU\\" + name, targetfullpath + "\\Capsule\\Windows")
                print(sourceFolder + "\\Combined\\WU\\" + name + " to " + targetFolder + "\\Capsule\\Windows" + " Copy succeeded.")
    # If Linux folder exist, copy files.
    if os.path.isdir(sourcefullpath+"\\Combined\\Linux"):
        for root,dirs,files in os.walk(sourcefullpath+"\\Combined\\Linux"):
            for name in files:
                copy(sourcefullpath + "\\Combined\\Linux\\" + name, targetfullpath + "\\Capsule\\Linux")
                print(sourceFolder + "\\Combined\\Linux\\" + name + " to " + targetFolder + "\\Capsule\\Linux" + " Copy succeeded.")
    #=======For G4 other Fv
    elif os.path.isfile(sourcefullpath + "\\Combined\\fwu.pfx"):
        for root,dirs,files in os.walk(sourcefullpath + "\\Combined"):
            for name in files:
                if name.find(".bin") != -1 or name.find(".inf") != -1:
                    copy(sourcefullpath + "\\Combined\\" + name, targetfullpath + "\\HPFWUPDREC")
                    print(sourceFolder + "\\Combined\\" + name + " to " + targetFolder + "\\HPFWUPDREC" + " Copy succeeded.")
                copy(sourcefullpath + "\\Combined\\" + name, targetfullpath + "\\Capsule")
                print(sourceFolder + "\\Combined\\" + name + " to " + targetFolder + "\\Capsule" + " Copy succeeded.")
    # Bin file copy to FPTW&Global.
    for name in os.listdir(sourcefullpath):# Bin file copy to FPTW&Global
        if (Platform_Flag(name) == "Q26") or (Platform_Flag(name) == "Q27") or \
            (Platform_Flag(name) == "R26") or (Platform_Flag(name) == "R24") or \
            (Platform_Flag(name) == "S25") or (Platform_Flag(name) == "S27") or (Platform_Flag(name) == "S29") or \
            (Platform_Flag(name) == "T25") or (Platform_Flag(name) == "T26") or (Platform_Flag(name) == "T27"): # 78 AMD 78 R26 78787878787878
            if name.find(NewVersion + "_16.bin") != -1:
                copy(sourcefullpath + name, targetfullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" +" Copy succeeded.")
            if name.find(NewVersion + "_32.bin") != -1:
                copy(sourcefullpath + name, targetfullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" +" Copy succeeded.")
        else:
            if name.find(NewVersion + ".bin") != -1:
                copy(sourcefullpath + name, targetfullpath + "\\AMDFLASH")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\AMDFLASH" + " Copy succeeded.")
        if name.find(NewVersion + ".bin") != -1:
            copy(sourcefullpath + name, targetfullpath + "\\Global\\BIOS")
            print(sourceFolder + "\\" + name + " to " + targetFolder + "\\Global\\BIOS" + " Copy succeeded.")
        if name.find(".xml") != -1:# Copy to XML
            if name.find(str(Platform_Flag(name))) != -1:
                copy(sourcefullpath + name, targetfullpath + "\\XML")
                print(sourceFolder + "\\" + name + " to " + targetFolder + "\\XML" + " Copy succeeded.")
        # For Smart flash copy *Pvt.bin.
        if name.find("Pvt.bin") != -1:
            copy(sourcefullpath + name, targetfullpath)
            print(sourceFolder + "\\" + name + " to " + targetFolder + " Copy succeeded.")
    print()


# Find Fv Folder, Add to Matchlist
def FindFvFolder(ProcessProjectList, NewVersion, NewBuildID):
    Matchlist = []
    for Process in ProcessProjectList:
        for Dir in os.listdir(".\\"):
            Match = "Fv_" + Process + "_" + NewVersion + "_"
            if not NewBuildID == "":
                Match = "Fv_" + Process + "_" + NewVersion + "_" + NewBuildID
            if Match in Dir:
                if not str(Dir).find(".zip") != -1 and not str(Dir).find(".7z") != -1:
                    Matchlist.append(Dir)
    return Matchlist


# Find Fv Zip file
def FindFvZip(ProcessProjectList, ProjectNameInfo, NewVersion, NewBuildID):
    Matchlist = []
    GiuhubPkgVersion = "FFFFFF"
    for i in range(len(ProcessProjectList)):
        for Dir in os.listdir(".\\"):
            #======For Old Fv
            Match = "Fv_" + ProcessProjectList[i] + "_" + NewVersion + "_"
            if not NewBuildID == "":
                Match = "Fv_" + ProcessProjectList[i] + "_" + NewVersion + "_" + NewBuildID
            if Match in Dir:
                if (str(Dir).find(".zip") != -1) or (str(Dir).find(".7z") != -1):
                    Matchlist.append(Dir)
            #======For New Github Fv
            if len(NewVersion) >= 6:
                GiuhubPkgVersion = "." + str(int(NewVersion[0:2])) + "." + str(int(NewVersion[2:4])) + "." + str(int(NewVersion[4:6]))
                GithubMatch = str(ProcessProjectList[i]).lower() + str(ProjectNameInfo[i]).lower() + GiuhubPkgVersion
            #======For U23 Github Fv project name mistake rename
            if ((str(ProcessProjectList[i]).lower() in Dir) or (str(ProcessProjectList[i]) in Dir)) and (GiuhubPkgVersion in Dir):
                TempDir = Dir.split(".")
                if TempDir[0] != str(ProcessProjectList[i]).lower() + str(ProjectNameInfo[i]).lower():
                    TempDir[0] = str(ProcessProjectList[i]).lower() + str(ProjectNameInfo[i]).lower()
                    TempDir = ".".join(TempDir)
                    os.rename(".\\" + Dir, ".\\" + TempDir)
                    Dir = TempDir
            #======Rename .nupkg to .zip
            if ((str(ProcessProjectList[i]).lower() in Dir) or (str(ProcessProjectList[i]) in Dir)) and (str(GiuhubPkgVersion) in Dir) and (str(Dir).find(".nupkg") != -1):
                NewDir = str(os.path.splitext(Dir)[0]) + ".zip" # Change package to zip
                os.rename(".\\" + Dir, ".\\" + NewDir)
                Matchlist.append(NewDir)
            if ((str(ProcessProjectList[i]).lower() in Dir) or (str(ProcessProjectList[i]) in Dir)) and (str(GiuhubPkgVersion) in Dir) and (str(Dir).find(".zip") != -1):
                Matchlist.append(Dir)
    return Matchlist


# Connect to FTP...
def Ftp_connect(Serverinfo):
    try:
        Ftp = FTP(Serverinfo["host"])
        Ftp.login(Serverinfo["username"], Serverinfo["pwd"])
        print("Connect to FTP......succeeded.")
        return Ftp
    except Exception as err:
        print("\nPlease check the network connection status.\nUnable to download fv Zip file from ftp now.\n")
        sys.exit()


def Ftp_get_file_list(line):
    file_arr = Ftp_get_filename(line)
    if file_arr[1] not in ['.', '..']:
        file_list.append(file_arr)


def Ftp_get_filename(line):
    pos = line.rfind(':')
    while line[pos] != ' ':
        pos += 1
    while line[pos] == ' ':
        pos += 1
    file_arr = [line[0], line[pos:]]
    return file_arr


# FTP callback 進度條.
def Ftp_callback(chunk, dst_file, lists, chunklist):
    max_arrow = 50  # 設定 '>' 有多少個
    listdate = "False"  # 檢查list的狀態
    listnum = 0
    tmp = 0
    if not hasattr(Ftp_callback, "progress_percent"):  # 初始化
        Ftp_callback.progress_percent = 0
    dst_file.write(chunk)  # 把 chunk 寫入到檔案中
    if len(chunklist) == 0:  # 假如chunklist中沒有資料，先寫入
        chunklist.append((str(os.getpid()), 0))  # 防止chunk二次增加，先寫入0
    for pid, data in chunklist:  # 尋找目前進程號，chunk疊加在寫回list
        if str(os.getpid()) == pid:
            data += len(chunk)
            listdate = "True"
            chunklist[listnum] = (str(os.getpid()), data)
        listnum += 1
    if listdate == "False":  # 假如沒有找到目前進程號，新增資料
        chunklist.append((str(os.getpid()), len(chunk)))
    if len(chunklist) != -1:
        for i in range(len(chunklist)):  # 把所有chunk相加
            tmp += chunklist[i][1]
        lists[1] = tmp
        tmp = lists[1] * 100 / lists[0]  # 得出目前要顯示的%數
        if Ftp_callback.progress_percent != tmp:
            Ftp_callback.progress_percent = tmp
            num_arrow = int(Ftp_callback.progress_percent / 2)  # 要顯示的">"
            num_line = max_arrow - num_arrow  # 要顯示的"-"
            #  要輸出的進度與%數，\r 是回到開頭在顯示，一直循環
            sys.stdout.write(
                "\r[" + ">" * num_arrow + "-" * num_line + "]" + "%.2f" % Ftp_callback.progress_percent + "%")
            if int(Ftp_callback.progress_percent) == 100:  # If 100%
                lists[1] = 0
                Ftp_callback.progress_percent = 0
            sys.stdout.flush()  # 刷新輸出


def Ftp_download_Test(localdir, remotedir, Test_Serverinfo):
    global file_list
    Ftp = Ftp_connect(Test_Serverinfo)  # FTP 連線
    try:
        Ftp.cwd(remotedir)
    except:
        print('Directory %s not exist,Continue...' % remotedir)
        return
    if not os.path.isdir(localdir):
        os.makedirs(localdir)
    print('Switch to directory %s' % Ftp.pwd())
    file_list = []
    Ftp.dir(Ftp_get_file_list)
    remotenames = file_list
    for item in remotenames:
        filetype = item[0]
        filename = item[1]
        local = os.path.join(localdir, filename)
        if filetype == 'd':
            Ftp_download_Test(local, remotedir + '/' + filename, Test_Serverinfo)
        elif filetype == '-':
            print('>>>>>>>>>>>>Download file %s ... ...' % local)
            file_handler = open(local, 'wb')
            Ftp.retrbinary('RETR %s' % (filename), file_handler.write)
            file_handler.close()
    Ftp.cwd('..')
    print('Return to upper directory %s' % Ftp.pwd())


# If Fv folder and .zip can't find, Download from HP FTP.
def Ftp_download(NProc, lists, chunklist, Serverinfo):
    Ftp = Ftp_connect(Serverinfo)  # FTP 連線
    Ftpfilelist = []
    Ftp_dir = "/ProductionReleaseDT/" + NProc[2] + "_" + NProc[0]
    if Serverinfo["type"] == "Test":
        Ftp_dir = "/Test Release/" + NProc[2]
    Ftp.cwd(Ftp_dir)
    Ftp.dir(Ftpfilelist.append)  # 寫入此資料夾下的所有檔案
    if len(NProc) == 4:
        Ftpfilelist = [File for File in Ftpfilelist if "Fv_" + NProc[2] + "_" + NProc[3] + "_0000_BuildJob_" in File]
        if len(Ftpfilelist) == 0:  # 如果找不到檔案
            print("Download " + "Fv_" + NProc[2] + "_" + NProc[3] + "_0000_BuildJob" + " Failed, Is not in the ftp site.")
            Ftp.quit()
            return False
    if len(NProc) == 5:
        Ftpfilelist = [File for File in Ftpfilelist if "Fv_" + NProc[2] + "_" + NProc[3] + "_" + NProc[4] + "_BuildJob_" in File]
        if len(Ftpfilelist) == 0:  # 如果找不到檔案
            print("Download " + "Fv_" + NProc[2] + "_" + NProc[3] + "_" + NProc[4] + "_BuildJob" + " Failed, Is not in the ftp site.")
            Ftp.quit()
            return False
    Fvfile = Ftpfilelist[0].split("Fv_")
    Fvfile = "Fv_" + Fvfile[1]  # 指定Fv檔案的全名
    Ftp.voidcmd("TYPE I")
    File_size = Ftp.size(Ftp_dir + "/" + Fvfile)  # 算出指定Fv檔案的大小
    if len(lists) == 0:
        lists.append(File_size)
        lists.append(0)
    else:
        lists[0] += File_size  # 計算出大小總和
    if not os.path.isfile(".\\" + Fvfile):  # 檢查Fv檔案是不是已經存在
        print("Start Download " + Fvfile + "\tfrom FTP, Please wait.....")
        f = open(Fvfile, "wb")  # 開啟Fv檔案準備寫入
        try:  # 開始下載&呼叫Ftp_callback載入進度
            sleep(0.5)
            Ftp.retrbinary("RETR %s/%s" % (Ftp_dir, Fvfile), lambda chunk: Ftp_callback(chunk, f, lists, chunklist))
        except error_perm:  # 如果出現錯誤
            print("Download " + Fvfile + " Failed, Is not in the ftp site.")
            f.close()
            Ftp.quit()
            return
        else:
            print()
            print("Download %s\tsuccessed\n" % Fvfile)
        f.close()  # 關閉FTP
        lists.append(Fvfile)  # 把下載完成的檔案加入list
    else:
        print(Fvfile + " have already exists.")  # 如果Fv檔案已經存在
    Ftp.quit()


# FTP 多進程下載
def Ftp_multi(NewProcPkgInfo, Production_Serverinfo, Test_Serverinfo):
    lists = Manager().list()  # 定義一個進程池可以共享的lists
    chunklist = Manager().list()
    po = Pool(10)  # 定義一個進程池，最大進程數
    Choose = InputStr("Download from ProductionSign key:1, TestSign key:2 :")
    if Choose == "1":
        Serverinfo = Production_Serverinfo
        for NProc in NewProcPkgInfo:
            # Pool.apply_async(要調用的目標,(傳遞給目標的參數元祖,))
            # 每次循環將會用空閒出來的子進程去調用目標
            po.apply_async(Ftp_download, (NProc, lists, chunklist, Serverinfo))
    else:
        Serverinfo = Test_Serverinfo
        for NProc in NewProcPkgInfo:
            po.apply_async(Ftp_download, (NProc, lists, chunklist, Serverinfo))
            ''' 因為test sign ftp也有放zip檔後就不用下載資料夾
            if len(NProc) == 4:
                localdir = '.' + r'\Fv_' + NProc[2] + '_' + NProc[3] + '_0000'  # 本地目錄
                remotedir = '/Test Release/' + NProc[2] + '/' + NProc[3][0: 2] + "." + NProc[3][2:4] + "." + NProc[3][4:6] + '_0000'  # 遠程目錄
            else:
                localdir = '.' + r'\Fv_' + NProc[2] + '_' + NProc[3] + '_' + NProc[4]  # 本地目錄
                remotedir = '/Test Release/' + NProc[2] + '/' + NProc[3][0: 2] + "." + NProc[3][2:4] + "." + NProc[3][4:6] + '_' + NProc[4]  # 遠程目錄
            po.apply_async(Ftp_download_Test, (localdir, remotedir, Test_Serverinfo))
            '''
    print("----start----")
    po.close()  # 關閉進程池，關閉後po不再接收新的請求
    po.join()  # 等待po中所有子進程執行完成，必須放在close語句之後
    print("-----end-----")
    return lists  # 返回進程池共用的lists


# Check New Release Pkg is OK?
def CheckPkg(NewProcPkgInfo, NewVersion):
    Check="False"
    for NProc in NewProcPkgInfo:
        Boardversion=NProc[2]+"_"+NProc[3]
        Path=".\\"+("_").join(NProc)
        if ((Platform_Flag(NewProcPkgInfo) == "Intel G3") or (Platform_Flag(NewProcPkgInfo) == "Intel G4")):
            if os.path.isdir(Path):
                if os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_12.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_32.bin"):
                    if os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_32.bin"):
                        if (os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".inf")) or \
                        (os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+"00.bin") or os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+"00.inf")) or \
                        (os.path.isfile(Path+"\\HPBIOSUPDREC\\"+Boardversion+".bin") or os.path.isfile(Path+"\\HPBIOSUPDREC\\"+Boardversion+".inf")):
                            if os.path.isfile(Path+"\\XML\\"+Boardversion+".xml"):
                                if not os.path.isfile(Path+"DCI.7z") or not os.path.isfile(Path+"DCI.zip"):
                                    Check="True"
                                else:print(Path+"DCI"+" can find.")
                            else:print(Path+"\\XML\\"+Boardversion+".xml"+" can't find.")
                        else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.");
                    else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
                else:print(Path+"\\FPTW\\"+Boardversion+"_12.bin"+" can't find.")
            else:print(Path+" can't find.")

        if ((Platform_Flag(NewProcPkgInfo) == "Intel G5") or (Platform_Flag(NewProcPkgInfo) == "Intel G6") or (Platform_Flag(NewProcPkgInfo) == "Intel G8") or (Platform_Flag(NewProcPkgInfo) == "Intel G9")or \
            (Platform_Flag(NewProcPkgInfo) == "Intel G10")):
            if os.path.isdir(Path):
                if ((os.path.isfile(Path+"\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)\\"+Boardversion+".bin") or os.path.isfile(Path+"\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)\\"+Boardversion+"00.bin")) or \
                (os.path.isfile(Path+"\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)\\"+Boardversion+".bin") or os.path.isfile(Path+"\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)\\"+Boardversion+"00.bin"))) and \
                ((Platform_Flag(NewProcPkgInfo) != "Intel G3") or (Platform_Flag(NewProcPkgInfo) != "Intel G4")):
                    if os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_12.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_32.bin"):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_32.bin"):
                            if (os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".inf")) or \
                            (os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+"00.bin") or os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+"00.inf")):
                                if os.path.isfile(Path+"\\XML\\"+Boardversion+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z") or not os.path.isfile(Path+"DCI.zip"):
                                        Check="True"
                                    else:print(Path+"DCI"+" can find.")
                                else:print(Path+"\\XML\\"+Boardversion+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
                    else:print(Path+"\\FPTW\\"+Boardversion+"_12.bin"+" can't find.")
                else:print(Path+"\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)\\"+Boardversion+".bin"+" can't find.");\
                    print(Path+"\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)\\"+Boardversion+".bin"+" can't find.");
            else:print(Path+" can't find.")

    if Check=="True":   print("New release Pkg made " + Fore.GREEN + "successfully.\n")
    else:   print("New release Pkg made " + Fore.RED + "failed.")


def CheckPkg_AMD(NewProcPkgInfo, NewVersion):
    Check="False"
    if (Platform_Flag(NewProcPkgInfo) == "R24") or (Platform_Flag(NewProcPkgInfo) == "Q26") or (Platform_Flag(NewProcPkgInfo) == "Q27"):
        for NProc in NewProcPkgInfo:
            Boardversion=NProc[0]+"_"+NewVersion
            NID=NProc[0]
            if (Platform_Flag(NewProcPkgInfo) == "R24"):
                Boardversion=NProc[1]+"_"+NewVersion
                NID=NProc[1]
            Path=".\\"+("_").join(NProc)
            if os.path.isdir(Path):
                if os.path.isfile(Path+"\\Capsule\\"+Boardversion+".bin") and os.path.isfile(Path+"\\Capsule\\"+Boardversion+".inf"):
                    if (os.path.isfile(Path+"\\AMDFLASH\\"+Boardversion+".bin") or os.path.isfile(Path+"\\AMDFLASH\\"+Boardversion+"_16.bin")) and \
                        os.path.isfile(Path+"\\Capsule\\"+NID.lower()+"_"+NewVersion+".cat"):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+".bin"):
                            if os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"):
                                if os.path.isfile(Path+"\\XML\\"+NID+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z"):
                                        Check="True"
                                    else:print(Path+"DCI.7z"+" can find.")
                                else:print(Path+"\\XML\\"+NID+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
                    else:print(Path+"\\AMDFLASH\\"+Boardversion+".bin or _16.bin"+" can't find.")
                else:print(Path+"\\Capsule\\"+Boardversion+".bin"+" can't find.");\
                    print(Path+"\\Capsule\\"+Boardversion+".inf"+" can't find.");print(Path+"\\Capsule\\"+NID.lower()+"_"+NewVersion+".cat"+" can't find.");
            else:print(Path+" can't find.")
    if (Platform_Flag(NewProcPkgInfo) == "R26") or (Platform_Flag(NewProcPkgInfo) == "S25"):# G5 and G6 AMD R26 AMD S25
        for NProc in NewProcPkgInfo:
            Boardversion=NProc[0]+"_"+NewVersion
            Path=".\\"+("_").join(NProc)
            if os.path.isdir(Path):
                if os.path.isfile(Path+"\\Capsule\\Windows\\"+Boardversion+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Boardversion+".inf"):
                    if (os.path.isfile(Path+"\\AMDFLASH\\"+Boardversion+".bin") or os.path.isfile(Path+"\\AMDFLASH\\"+Boardversion+"_16.bin")) and \
                        os.path.isfile(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat"):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+".bin"):
                            if os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"):
                                if os.path.isfile(Path+"\\XML\\"+NProc[0]+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z"):
                                        Check="True"
                                    else:print(Path+"DCI.7z"+" can find.")
                                else:print(Path+"\\XML\\"+NProc[0]+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
                    else:print(Path+"\\AMDFLASH\\"+Boardversion+".bin or _16.bin"+" can't find.")
                else:print(Path+"\\Capsule\\Windows\\"+Boardversion+".bin"+" can't find.");\
                    print(Path+"\\Capsule\\Windows\\"+Boardversion+".inf"+" can't find.");print(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat"+" can't find.");
            else:print(Path+" can't find.")
    if (Platform_Flag(NewProcPkgInfo) == "S27") or (Platform_Flag(NewProcPkgInfo) == "S29") or (Platform_Flag(NewProcPkgInfo) == "T25") or \
        (Platform_Flag(NewProcPkgInfo) == "T27") or (Platform_Flag(NewProcPkgInfo) == "T26"):# G5 and G6 AMD R26 AMD S25
        for NProc in NewProcPkgInfo:
            Boardversion=NProc[0]+"_"+NewVersion
            Path=".\\"+("_").join(NProc)
            if os.path.isdir(Path):
                if os.path.isfile(Path+"\\Capsule\\Windows\\"+Boardversion+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Boardversion+".inf"):
                    if (os.path.isfile(Path+"\\AMDFLASH\\"+Boardversion+".bin") or os.path.isfile(Path+"\\AMDFLASH\\"+Boardversion+"_32.bin")) and \
                        os.path.isfile(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat"):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+".bin"):
                            if os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"):
                                if os.path.isfile(Path+"\\XML\\"+NProc[0]+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z"):
                                        Check="True"
                                    else:print(Path+"DCI.7z"+" can find.")
                                else:print(Path+"\\XML\\"+NProc[0]+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
                    else:print(Path+"\\AMDFLASH\\"+Boardversion+".bin or _32.bin"+" can't find.")
                else:print(Path+"\\Capsule\\Windows\\"+Boardversion+".bin"+" can't find.");\
                    print(Path+"\\Capsule\\Windows\\"+Boardversion+".inf"+" can't find.");print(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat"+" can't find.");
            else:print(Path+" can't find.")
    if Check=="True":   print("New release Pkg made "+Fore.GREEN+"successfully.\n")
    else:   print("New release Pkg made "+Fore.RED+"failed.")


# Return[ver,mtime,size]
def GetFileInfo(filepath):
    try:
        file_info = []
        #getfileversion
        ver_parser = Dispatch('Scripting.FileSystemObject')
        ver = ver_parser.GetFileVersion(filepath)
        if ver == 'No Version Information Available':
            ver = None
        file_info.append(ver)
        #getfilemtime
        tmpTime = localtime(os.stat(filepath).st_mtime)
        file_info.append(strftime('%Y/%m/%d', tmpTime))
        #getfilesize
        size = os.path.getsize(filepath)
        file_info.append(size)
        return file_info
    except:
        return None


# Return[Name, Category, Version, Path, Notes]
def ReadToolVersionTable(tablepath):
    app = xw.App(visible = False,add_book = False)
    app.display_alerts = False
    app.screen_updating = False
    toolinfo = []
    fn = tablepath
    wb = app.books.open(fn)
    if wb.sheets[0].name == "ENT18 BIOS&ME tool list":
        ToolList = wb.sheets["ENT18 BIOS&ME tool list"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,12
        for i in range(3, int(UsRange[4][:2])+1): # Lines ex:3 to 13
            ##['<Range [ToolVersion.xlsx]ENT18 BIOS&ME tool list', 'A', '1:', 'E', '12>']
            ###### I suggeset [:2] changing to [:-1] 
            temp=[]
            for j in range(1, ord(UsRange[3])-64+1): # Columns ex:1 to 6
                temp.append(ToolList.range((i, j)).value)
            toolinfo.append(temp)
    elif wb.sheets[0].name == "ENT19 BIOS": #AMD   don't have ME
        ToolList = wb.sheets["ENT19 BIOS"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,7
        #print(UsRange)                                # ['<Range [ToolVersion.xlsx]ENT19 BIOS!', 'A', '1:', 'E', '7>']
        #onlyNum = sub("[0-9]+","",UsRange[4][:-1]) # 7
        for i in range(2, int(UsRange[4][:-1])+1): # Lines ex:2 to 7
            temp=[]
            for j in range(1, ord(UsRange[3])-64+1): # Columns ex:1 to 6
                temp.append(ToolList.range((i, j)).value)
            toolinfo.append(temp)
    wb.save()
    wb.close()
    app.quit()
    return toolinfo


# Set version info to table
def SetToolVersionTable(tablepath, name, fileinfo_temp):
    app = xw.App(visible = False,add_book = False)
    app.display_alerts = False
    app.screen_updating = False
    fn = tablepath
    wb = app.books.open(fn)
    if wb.sheets[0].name == "ENT18 BIOS&ME tool list":
        ToolList = wb.sheets["ENT18 BIOS&ME tool list"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,12
        for i in range(3,int(UsRange[4][:2])+1): # Lines ex:3 to 13
            # I suggeset [:2] changing to [:-1] too 
            if ToolList.range(i,1).value == name:
                ToolList.range(i,3).value = str(fileinfo_temp)
    # AMD
    elif wb.sheets[0].name == "ENT19 BIOS":
        ToolList = wb.sheets["ENT19 BIOS"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,7
        for i in range(2,int(UsRange[4][:-1])+1): # Lines ex:2 to 7
            if ToolList.range(i,1).value == name:
                ToolList.range(i,3).value = str(fileinfo_temp)
    wb.save()
    wb.close()
    app.quit()


def CompareInfo(NProc, name, ver, path, Toolversiontablepath):
    PkgPath = ".\\" + ("_").join(NProc) + "\\"
    FilesPath = {"MEInfoWin64.exe":          PkgPath + "METools\\MEInfo\\WINDOWS64\\",
                "MEManufWin64.exe":         PkgPath + "METools\\MEManuf\\WINDOWS64\\",
                "FPTW.exe":                 PkgPath + "FPTW\\",
                "FWUpdLcl64.exe":           PkgPath + "FactoryUtility\\", # Zip
                "EEUPDATEW64e.exe":         PkgPath + "FPTW\\",
                "BiosConfigUtility64.exe":  PkgPath + "FPTW\\",
                "BiosConfigUtility.exe":    PkgPath + "AMDFLASH\\NeverLock\\",  # AMD
                "HpFirmwareUpdRec64.exe":   PkgPath + "HPFWUPDREC\\",
                "HpFirmwareUpdRec.exe":     PkgPath + "HPFWUPDREC\\",           # AMD
                "Buff2.efi":                PkgPath + "FPTW\\",# Skip
                "ElectronicLabelUpdate.efi":PkgPath + "FactoryUtility\\", # Zip # Skip
                "wmiTool64.exe":            PkgPath + "FactoryUtility\\"} # Zip # Skip
    if name in FilesPath:
        if name == "FWUpdLcl64.exe" or name == "ElectronicLabelUpdate.efi" or name == "wmiTool64.exe": #Files in zip
            if name == "FWUpdLcl64.exe":
                if os.path.isfile(PkgPath + "FactoryUtility\\FWUpdate.zip"):
                    UnZip(PkgPath + "FactoryUtility\\FWUpdate.zip")
                fileinfo_temp = GetFileInfo(PkgPath + "FactoryUtility\\FWUpdate\\Windows64\\FWUpdLcl64.exe")
                rmtree(PkgPath + "FactoryUtility\\FWUpdate")
                if not fileinfo_temp[0] == ver:
                    SetToolVersionTable(Toolversiontablepath, name, fileinfo_temp[0])
                    print("Set " + str(name) + " Version: " + str(ver) + " to " + str(fileinfo_temp[0]))
            else:
                pass
        elif name == "MEInfoWin64.exe" or name == "MEManufWin64.exe" or name == "FPTW.exe" or name == "EEUPDATEW64e.exe" or\
            name == "BiosConfigUtility64.exe" or name == "HpFirmwareUpdRec64.exe" or name == "BiosConfigUtility.exe" or\
            name == "HpFirmwareUpdRec.exe" : # add AMD
            if os.path.isfile(FilesPath[name] + name):
                fileinfo_temp = GetFileInfo(FilesPath[name] + name)
                if not fileinfo_temp[0] == ver:
                    SetToolVersionTable(Toolversiontablepath, name, fileinfo_temp[0])
                    print("Set " + str(name) + " Version: " + str(ver) + " to " + str(fileinfo_temp[0]))
        elif name == None or name == "Buff2.efi": #Skip compare
            pass
        else:
            pass

def ChangeVersionInfo(verinfo):
    str(verinfo).strip()
    str(verinfo).lower()
    if str(verinfo).find("v") >= 0:
        ver= verinfo.lstrip("v")
    else:
        ver = verinfo
    return ver


def ChangeDataInfo(dateinfo):
    str(dateinfo).strip()
    date = dateinfo.strftime("%Y/%m/%d")
    return date


def CheckBiosBuildDate(Matchfolderlist):
    BiosBuildDate = {}
    for Fv in Matchfolderlist:
        time = ""
        if os.path.isdir(".\\" + Fv):
            Path = ".\\" + Fv + "\\Combined\\"
            for root,dirs,files in os.walk(Path):
                for name in files:
                    if ".inf" in name:
                        pattern = r"\d+/\d+/\d+"
                        File = open(root + "\\" + name, "r+")
                        Strlist = File.read()
                        searchObj = search(pattern, Strlist)
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


def GetMrcVersion(Matchfolderlist):
    for Fv in Matchfolderlist:
        MrcVersion = ""
        if (Platform_Flag(Fv) == "Intel G5"): #Block Intel G5 for MRC Version error.
            return (MrcVersion)
        if (Platform_Flag(Fv) == "Intel G6"):
            MRC_Offset = 4
        else:
            MRC_Offset = 0
        Path = ".\\" + Fv
        if os.path.isdir(Path):
            for root, dirs, files in os.walk(Path):
                for name in files:
                    if name.find("_32.bin") != -1:
                        logging.debug("find BIOS binary: " + name)
                        MRCVer_Str = b'MRCVER_'
                        with open(root + "\\" + name, "rb") as BinaryFile:
                            BinaryData = BinaryFile.read()
                            MRCVer_Str_index = BinaryData.find(MRCVer_Str)
                        if MRCVer_Str_index == -1:
                            print(Fore.RED + 'Can not find MRC Version')
                        else:
                            MRCVer_End_index = MRCVer_Str_index + \
                                len(MRCVer_Str) + MRC_Offset + 4
                            VersionData = BinaryData[MRCVer_Str_index +
                                                     len(MRCVer_Str) + MRC_Offset:MRCVer_End_index]
                            byte_sequence = struct.unpack('>4B', VersionData)
                            MrcVersion = '.'.join(map(str, byte_sequence))
                        break
                    else:
                        logging.debug('Can not find BIOS binary')
        if MrcVersion == "":
            logging.debug("Get MRC Version fail.")
    return (MrcVersion)


def PrintBiosBuildDate(Matchfolderlist, BiosBuildDate):
    for Fv in Matchfolderlist:
        print(Fv.split("_")[1]+"_"+Fv.split("_")[2]+" Build Date:"+BiosBuildDate[Fv.split("_")[1]])


def CheckFileChecksum(Matchfolderlist, NewVersion):
    try:
        BiosFileChecksum = {}
        for NProc in Matchfolderlist:
            path = ".\\" + NProc
            if (Platform_Flag(NProc) == "Intel G3") or (Platform_Flag(NProc) == "Intel G4") or (Platform_Flag(NProc) == "Intel G5") or \
                (Platform_Flag(NProc) == "Intel G6") or (Platform_Flag(NProc) == "Intel G8") or (Platform_Flag(NProc) == "Intel G9") or \
                (Platform_Flag(NProc) == "Intel G10"):
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
                            if name.find("_16.bin") != -1:
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
        if (Platform_Flag(NProc) == "Intel G3") or (Platform_Flag(NProc) == "Intel G4") or (Platform_Flag(NProc) == "Intel G5") or \
            (Platform_Flag(NProc) == "Intel G6") or (Platform_Flag(NProc) == "Intel G8") or (Platform_Flag(NProc) == "Intel G9") or \
            (Platform_Flag(NProc) == "Intel G10"):
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


def CheckMEVersion(NProc, Matchfolderlist):
    Version = "11.0.11.1111" # Default version
    logging.debug('CheckMEVersion Start.')
    for Fv in Matchfolderlist:
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
                Strlist = File.read()
                searchObj = search(pattern, Strlist)
                if searchObj != None:
                    Version = searchObj.group(0)[25:27] + ".0." + searchObj.group(0)[28:30] + "." + searchObj.group(0)[31:33] + searchObj.group(0)[34:36]
                    File.close()
                logging.debug('ME_Version2:' + Version)
                return Version
            else:
                return Version


#Not Used now.
def SetReleaseNoteVersionValue(Version):
    RevisionV1075        = {'Revision':                  'A8'}
    IntelProjectPNV1075  = {'BIOS VERSION Value':        'C6',
                            'BIOS PART NUMBER Value':    'C7',
                            'ME VERSION Value':          'C11',
                            'ME PART NUMBER':            'C12'}
    IntelInfoV1075       = {'Folder Path':               'A26',
                            'ODM FTP':                   'A27',
                            'Folder Path':               'A28',
                            'Folder Path Value':         'B26',
                            'ODM FTP Value':             'B27',
                            'Folder Path Value':         'B28'}
    IntelHistoryV1075    = {'System BIOS Version':       'A10',
                            'Target EE phase (DB/SI/PV)':'A11',
                            'Build Date':                'A12',
                            'CHECKSUM':                  'A13',
                            'System BIOS Version Value': 'B10',
                            'Build Date Value':          'B12',
                            'CHECKSUM Value':            'B13',
                            'ME Firmware':               'B35'}
    IntelHowToFlashV1075 = {'BIOS Flash: From -> To':    'A18'}


def ModifyReleaseNote(NProc, ReleaseFileName, BiosBuildDate, BiosBinaryChecksum, NewVersion, NewBuildID, BiosMrcVersion, Matchfolderlist):
    print("Platform ReleaseNote Modify...")
    app = xw.App(visible = False,add_book = False)
    app.display_alerts = False
    app.screen_updating = False
    filepath = ReleaseFileName
    wb = app.books.open(filepath)
    if "v1.08" in wb.sheets['Revision'].range('A8').value or \
    "v1.07" in wb.sheets['Revision'].range('A8').value or \
    "v1.06" in wb.sheets['Revision'].range('A8').value:
        try:
            #======If Intel DM G5 and late
            if (Platform_Flag(ReleaseFileName) == "Intel G5") or (Platform_Flag(ReleaseFileName) == "Intel G6") or \
                (Platform_Flag(ReleaseFileName) == "Intel G8") or (Platform_Flag(ReleaseFileName) == "Intel G9") or \
                (Platform_Flag(ReleaseFileName) == "Intel G10"):
                logging.debug('If Intel DM G5 and late')
                MEVersion = CheckMEVersion(NProc, Matchfolderlist) # ex. 14.0.21.7227
                IntelHistory = wb.sheets['IntelPlatformHistory']
                IntelProjectPN = wb.sheets['IntelProjectPN']
                IntelInfo = wb.sheets['IntelPlatformInfo']
                IntelHowToFlash = wb.sheets['IntelPlatformHowToFlash']
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
                for a in range(9, 30):
                    if IntelHistory.range('A'+str(a)).value == 'System BIOS Version' and \
                    IntelHistory.range('A'+str(a+1)).value == 'Target EE phase (DB/SI/PV)' and \
                    IntelHistory.range('A'+str(a+2)).value == 'Build Date':
                        for b in range(3, 30):
                            if IntelProjectPN.range('B'+str(b)).value == 'VERSION' and \
                            IntelProjectPN.range('B'+str(b+1)).value == 'PART NUMBER' and \
                            IntelProjectPN.range('B'+str(b+5)).value == 'VERSION':
                                if IntelHistory.range('A'+str(a+3)).value == 'BIOS Build Version':
                                    logging.debug('IntelHistory(a+3) = BIOS Build Version')
                                    if (NewBuildID == "" or NewBuildID == "0000"):
                                        IntelHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
                                        IntelProjectPN.range('C'+str(b)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
                                        IntelHistory.range('B'+str(a+3)).value = "0000"
                                    else:
                                        IntelHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID
                                        IntelProjectPN.range('C'+str(b)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID
                                        IntelHistory.range('B'+str(a+3)).value = NewBuildID
                                    logging.debug('IntelHistory(a+4) = BIOS Checksum')
                                    IntelHistory.range('B'+str(a+4)).value = "0x" + BiosBinaryChecksum[NProc[2]].upper() #CHECK SUM
                                    logging.debug('IntelHistory(a+6) = MRC')
                                    if BiosMrcVersion !="":
                                        IntelHistory.range('B'+str(a+6)).value = "" + str(BiosMrcVersion)  # MRC VERSION
                                        if(IntelHistory.range('B'+str(a+6)).value != IntelHistory.range('C'+str(a+6)).value):
                                            logging.debug('MRC Version change')
                                            IntelHistory.range('B'+str(a+6)).api.Font.Color = 0x00B050
                                elif IntelHistory.range('A'+str(a+3)).value == 'CHECKSUM':
                                    logging.debug('IntelHistory(a+3) = CHECKSUM')
                                    if (NewBuildID == "" or NewBuildID == "0000"):
                                        IntelHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
                                        IntelProjectPN.range('C'+str(b)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6]
                                    else:
                                        IntelHistory.range('B'+str(a)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID
                                        IntelProjectPN.range('C'+str(b)).value = NewVersion[0:2] + "." + NewVersion[2:4] + "." + NewVersion[4:6] + "_" + NewBuildID
                                    logging.debug('IntelHistory(a+4) = BIOS Checksum')
                                    IntelHistory.range('B'+str(a+3)).value = "0x" + BiosBinaryChecksum[NProc[2]].upper() #CHECK SUM
                                    if BiosMrcVersion !="":
                                        logging.debug('IntelHistory(a+5) = MRC')
                                        IntelHistory.range('B'+str(a+5)).value = "" + str(BiosMrcVersion)  # MRC VERSION
                                        if(IntelHistory.range('B'+str(a+5)).value != IntelHistory.range('C'+str(a+5)).value):
                                            logging.debug('MRC Version change')
                                            IntelHistory.range('B'+str(a+5)).api.Font.Color = 0x00B050
                                else:
                                    check = "fail"
                                    break
                                IntelProjectPN.range('C'+str(b+1)).value = "P00000-000" #PART NUMBER
                                IntelHistory.range('B'+str(a+2)).value = BiosBuildDate[ReleaseFileName.split("_")[2]] #BUILD DATE
                                logging.debug('Version fill finish.')
                                check = "pass"
                                break
                        if check == "pass":
                            break
                if not check == "pass":
                    print("Can't find ['System BIOS Version', 'Target EE phase (DB/SI/PV)', 'Build Date', 'CHECKSUM']")
                #======ME Version
                logging.debug('ME Version')
                if not MEVersion == "11.0.11.1111":
                    pattern = r'[0-9]+[\.][0-9]+[\.][0-9]+[\.]\d{4}'
                    check = ""
                    for a in range(25, 50):
                        if IntelHistory.range('A'+str(a)).value == 'ME Firmware':
                            OldMEVersionStr = IntelHistory.range('B'+str(a)).value
                            logging.debug("OldMEVersionStr = " + OldMEVersionStr)
                            NewMEVersionStr = sub(pattern, MEVersion, OldMEVersionStr)
                            logging.debug("NewMEVersionStr = " + NewMEVersionStr)
                            searchObj = search(pattern, OldMEVersionStr)
                            if searchObj != None:
                                OldMEVersion = searchObj.group(0)
                            else:
                                break
                            IntelHistory.range('B'+str(a)).value = NewMEVersionStr
                            logging.debug('IntelHistory.range(B+str(a)).value = ' + IntelHistory.range('B'+str(a)).value)
                            for b in range(3, 30):
                                if IntelProjectPN.range('B'+str(b)).value == 'VERSION' and \
                                    IntelProjectPN.range('B'+str(b+1)).value == 'PART NUMBER' and \
                                    IntelProjectPN.range('B'+str(b+5)).value == 'VERSION':
                                    IntelProjectPN.range('C'+str(b+5)).value = MEVersion
                                    if str(MEVersion).strip() != str(OldMEVersion).strip():
                                        IntelProjectPN.range('C'+str(b+6)).value = "P00000-000" #PART NUMBER
                                        print("Set ME FW Part Number to P00000-000.")
                                    logging.debug('ME fill finish.')
                                    check = "pass"
                                    break
                            logging.debug('ME fill finish2.')
                        if check == "pass":
                            break
                if not check == "pass":
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
                if not check == "pass":
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
                if not check == "pass":
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
            #======If AMD G5 DM
            elif (Platform_Flag(ReleaseFileName) == "R26") or (Platform_Flag(ReleaseFileName) == "R24") or (Platform_Flag(ReleaseFileName) == "S25") or \
                    (Platform_Flag(ReleaseFileName) == "T26") or (Platform_Flag(ReleaseFileName) == "T27") or (Platform_Flag(ReleaseFileName) == "T25") or \
                    (Platform_Flag(ReleaseFileName) == "S27") or (Platform_Flag(ReleaseFileName) == "S29"):
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
                        AMDHistory.range('B'+str(a+3)).value = "0x" + BiosBinaryChecksum[NProc[2]].upper()
                        logging.debug('Version fill finish.')
                        check = "pass"
                        break
                if not check == "pass":
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
                if not check == "pass":
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
                if not check == "pass":
                    print("Can't find ['BIOS Flash: From -> To ']")
            wb.save()
            wb.close()
            app.quit()
            print("Platform ReleaseNote Modify " + Fore.GREEN + "succeeded.\n")
        except:
            wb.close()
            app.quit()
            print("ReleaseNote Modify " + Fore.RED + "Failed!\n")
            return 0
    else:
        Old_ReleaseNoteVersion = wb.sheets['Revision'].range('A8').value
        wb.close()
        app.quit()
        print("ReleaseNote Version is " + Old_ReleaseNoteVersion + " Not in 1.06~1.08 Version, Modify Skip.\n")
