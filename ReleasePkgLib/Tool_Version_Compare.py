import os
from shutil import rmtree
from win32com.client import Dispatch
from time import localtime, strftime
import xlwings as xw

from ReleasePkgLib import *


# Return[ver,mtime,size]
def GetFileInfo(filepath):
    try:
        file_info = []
        #get_file_version
        ver_parser = Dispatch('Scripting.FileSystemObject')
        ver = ver_parser.GetFileVersion(filepath)
        if ver == 'No Version Information Available':
            ver = None
        file_info.append(ver)
        #get_file_time
        tmpTime = localtime(os.stat(filepath).st_mtime)
        file_info.append(strftime('%Y/%m/%d', tmpTime))
        #get_file_size
        size = os.path.getsize(filepath)
        file_info.append(size)
        return file_info
    except:
        return None


# Return[Name, Category, Version, Path, Notes]
def ReadToolVersionTable(table_path):
    app = xw.App(visible = False,add_book = False)
    app.display_alerts = False
    app.screen_updating = False
    tool_info = []
    fn = table_path
    wb = app.books.open(fn)
    if wb.sheets[0].name == "ENT18 BIOS&ME tool list":
        ToolList = wb.sheets["ENT18 BIOS&ME tool list"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,12
        for i in range(3, int(UsRange[4][:2])+1): # Lines ex:3 to 13
            ##['<Range [ToolVersion.xlsx]ENT18 BIOS&ME tool list', 'A', '1:', 'E', '12>']
            ###### I suggest [:2] changing to [:-1]
            temp=[]
            for j in range(1, ord(UsRange[3])-64+1): # Columns ex:1 to 6
                temp.append(ToolList.range((i, j)).value)
            tool_info.append(temp)
    elif wb.sheets[0].name == "ENT19 BIOS": #AMD   don't have ME
        ToolList = wb.sheets["ENT19 BIOS"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,7
        #print(UsRange)                                # ['<Range [ToolVersion.xlsx]ENT19 BIOS!', 'A', '1:', 'E', '7>']
        #onlyNum = sub("[0-9]+","",UsRange[4][:-1]) # 7
        for i in range(2, int(UsRange[4][:-1])+1): # Lines ex:2 to 7
            temp=[]
            for j in range(1, ord(UsRange[3])-64+1): # Columns ex:1 to 6
                temp.append(ToolList.range((i, j)).value)
            tool_info.append(temp)
    wb.save()
    wb.close()
    app.quit()
    return tool_info


# Set version info to table
def SetToolVersionTable(table_path, name, file_info_temp):
    app = xw.App(visible = False,add_book = False)
    app.display_alerts = False
    app.screen_updating = False
    fn = table_path
    wb = app.books.open(fn)
    if wb.sheets[0].name == "ENT18 BIOS&ME tool list":
        ToolList = wb.sheets["ENT18 BIOS&ME tool list"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,12
        for i in range(3,int(UsRange[4][:2])+1): # Lines ex:3 to 13
            # I suggest [:2] changing to [:-1] too
            if ToolList.range(i,1).value == name:
                ToolList.range(i,3).value = str(file_info_temp)
    # AMD
    elif wb.sheets[0].name == "ENT19 BIOS":
        ToolList = wb.sheets["ENT19 BIOS"]
        UsRange = str(ToolList.used_range).split('$') # get file date range, ex:A,1,E,7
        for i in range(2,int(UsRange[4][:-1])+1): # Lines ex:2 to 7
            if ToolList.range(i,1).value == name:
                ToolList.range(i,3).value = str(file_info_temp)
    wb.save()
    wb.close()
    app.quit()


def CompareInfo(NProc, name, ver, path, Tool_version_table_path):
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
                file_info_temp = GetFileInfo(PkgPath + "FactoryUtility\\FWUpdate\\Windows64\\FWUpdLcl64.exe")
                rmtree(PkgPath + "FactoryUtility\\FWUpdate")
                if not file_info_temp[0] == ver:
                    SetToolVersionTable(Tool_version_table_path, name, file_info_temp[0])
                    print("Set " + str(name) + " Version: " + str(ver) + " to " + str(file_info_temp[0]))
            else:
                pass
        elif name == "MEInfoWin64.exe" or name == "MEManufWin64.exe" or name == "FPTW.exe" or name == "EEUPDATEW64e.exe" or\
            name == "BiosConfigUtility64.exe" or name == "HpFirmwareUpdRec64.exe" or name == "BiosConfigUtility.exe" or\
            name == "HpFirmwareUpdRec.exe" : # add AMD
            if os.path.isfile(FilesPath[name] + name):
                file_info_temp = GetFileInfo(FilesPath[name] + name)
                if not file_info_temp[0] == ver:
                    SetToolVersionTable(Tool_version_table_path, name, file_info_temp[0])
                    print("Set " + str(name) + " Version: " + str(ver) + " to " + str(file_info_temp[0]))
        elif name == None or name == "Buff2.efi": #Skip compare
            pass
        else:
            pass

def ChangeVersionInfo(ver_info):
    str(ver_info).strip()
    str(ver_info).lower()
    if str(ver_info).find("v") >= 0:
        ver= ver_info.lstrip("v")
    else:
        ver = ver_info
    return ver


def ChangeDataInfo(date_info):
    str(date_info).strip()
    date = date_info.strftime("%Y/%m/%d")
    return date