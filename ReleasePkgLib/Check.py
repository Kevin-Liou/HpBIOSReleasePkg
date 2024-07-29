import os
from colorama import Fore

from ReleasePkgLib import *
from .Platform import *


# Check New Release Pkg is OK?
def CheckPkg(NewProcPkgInfo):
    Check="False"
    for NProc in NewProcPkgInfo:
        Board_version=NProc[2]+"_"+NProc[3]
        Path=".\\"+("_").join(NProc)
        # Check Intel Project G3 G4
        if Platform_Flag(NewProcPkgInfo) in Intel_Platforms_G3G4:
            if os.path.isdir(Path):
                if os.path.isfile(Path+"\\FPTW\\"+Board_version+"_12.bin") or os.path.isfile(Path+"\\FPTW\\"+Board_version+"_16.bin") or os.path.isfile(Path+"\\FPTW\\"+Board_version+"_32.bin"):
                    if os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+"_16.bin") or os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+"_32.bin"):
                        if (os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".inf")) or \
                        (os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+"00.bin") or os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+"00.inf")) or \
                        (os.path.isfile(Path+"\\HPBIOSUPDREC\\"+Board_version+".bin") or os.path.isfile(Path+"\\HPBIOSUPDREC\\"+Board_version+".inf")):
                            if os.path.isfile(Path+"\\XML\\"+Board_version+".xml"):
                                if not os.path.isfile(Path+"DCI.7z") or not os.path.isfile(Path+"DCI.zip"):
                                    Check="True"
                                else:print(Path+"DCI"+" can find.")
                            else:print(Path+"\\XML\\"+Board_version+".xml"+" can't find.")
                        else:print(Path+"\\HPFWUPDREC\\"+Board_version+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Board_version+".inf"+" can't find.");
                    else:print(Path+"\\Global\\BIOS\\"+Board_version+".bin"+" can't find.")
                else:print(Path+"\\FPTW\\"+Board_version+"_12.bin"+" can't find.")
            else:print(Path+" can't find.")
        # Check Intel Project G5 G6 G8 G9 G10
        if Platform_Flag(NewProcPkgInfo) in Intel_Platforms_G5later:
            if os.path.isdir(Path):
                if ((os.path.isfile(Path+"\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)\\"+Board_version+".bin") or os.path.isfile(Path+"\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)\\"+Board_version+"00.bin")) or \
                (os.path.isfile(Path+"\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)\\"+Board_version+".bin") or os.path.isfile(Path+"\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)\\"+Board_version+"00.bin"))) and \
                ((Platform_Flag(NewProcPkgInfo) != "Intel G3") or (Platform_Flag(NewProcPkgInfo) != "Intel G4")):
                    if os.path.isfile(Path+"\\FPTW\\"+Board_version+"_12.bin") or os.path.isfile(Path+"\\FPTW\\"+Board_version+"_16.bin") or os.path.isfile(Path+"\\FPTW\\"+Board_version+"_32.bin"):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+"_16.bin") or os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+"_32.bin"):
                            if (os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".inf")) or \
                            (os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+"00.bin") or os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+"00.inf")):
                                if os.path.isfile(Path+"\\XML\\"+Board_version+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z") or not os.path.isfile(Path+"DCI.zip"):
                                        Check="True"
                                    else:print(Path+"DCI"+" can find.")
                                else:print(Path+"\\XML\\"+Board_version+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Board_version+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Board_version+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Board_version+".bin"+" can't find.")
                    else:print(Path+"\\FPTW\\"+Board_version+"_12.bin"+" can't find.")
                else:print(Path+"\\Capsule\\Linux\\Combined FW Image (BIOS, ME, PD)\\"+Board_version+".bin"+" can't find.");\
                    print(Path+"\\Capsule\\Windows\\Combined FW Image (BIOS, ME, PD)\\"+Board_version+".bin"+" can't find.");
            else:print(Path+" can't find.")

    if Check=="True":   print("New release Pkg made " + Fore.GREEN + "successfully.\n")
    else:   print("New release Pkg made " + Fore.RED + "failed.")


def CheckPkg_AMD(NewProcPkgInfo, NewVersion, NewBuildID):
    Check="False"
    # Check AMD Project Q26 Q27 R24
    if  (Platform_Flag(NewProcPkgInfo) == "Q26") or (Platform_Flag(NewProcPkgInfo) == "Q27") or (Platform_Flag(NewProcPkgInfo) == "R24"):
        for NProc in NewProcPkgInfo:
            Board_version=NProc[0]+"_"+NewVersion
            NID=NProc[0]
            if (Platform_Flag(NewProcPkgInfo) == "R24"):
                Board_version=NProc[1]+"_"+NewVersion
                NID=NProc[1]
            Path=".\\"+("_").join(NProc)
            if os.path.isdir(Path):
                if (os.path.isfile(Path+"\\Capsule\\"+Board_version+".bin") and os.path.isfile(Path+"\\Capsule\\"+Board_version+".inf"))\
                or(os.path.isfile(Path+"\\Capsule\\"+Board_version+NewBuildID[-2:]+".bin") and os.path.isfile(Path+"\\Capsule\\"+Board_version+NewBuildID[-2:]+".inf")) :
                    if (os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+".bin") or os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+"_16.bin")) and \
                        os.path.isfile(Path+"\\Capsule\\"+NID.lower()+"_"+NewVersion+".cat"):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+".bin"):
                            if os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".inf"):
                                if os.path.isfile(Path+"\\XML\\"+NID+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z"):
                                        Check="True"
                                    else:print(Path+"DCI.7z"+" can find.")
                                else:print(Path+"\\XML\\"+NID+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Board_version+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Board_version+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Board_version+".bin"+" can't find.")
                    else:print(Path+"\\AMDFLASH\\"+Board_version+".bin or _16.bin"+" can't find.")
                else:print(Path+"\\Capsule\\"+Board_version+".bin"+" can't find.");\
                    print(Path+"\\Capsule\\"+Board_version+".inf"+" can't find.");print(Path+"\\Capsule\\"+NID.lower()+"_"+NewVersion+".cat"+" can't find.");
            else:print(Path+" can't find.")
    # Check AMD Project R26 S25
    if (Platform_Flag(NewProcPkgInfo) == "R26") or (Platform_Flag(NewProcPkgInfo) == "S25"):
        for NProc in NewProcPkgInfo:
            Board_version=NProc[0]+"_"+NewVersion
            Path=".\\"+("_").join(NProc)
            if os.path.isdir(Path):
                if (os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+".inf"))\
                or(os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+NewBuildID[-2:]+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+NewBuildID[-2:]+".inf")) \
		        or(os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"00.bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"00.inf")) :
                    if (os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+".bin") or os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+"_16.bin")) and \
                        os.path.isfile(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat")\
			            and \
			            (os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"_"+NewVersion+".cat")\
                        or os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"00.cat")):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+".bin"):
                            if os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".inf"):
                                if os.path.isfile(Path+"\\XML\\"+NProc[0]+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z"):
                                        Check="True"
                                    else:print(Path+"DCI.7z"+" can find.")
                                else:print(Path+"\\XML\\"+NProc[0]+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Board_version+"(00).bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Board_version+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Board_version+".bin"+" can't find.")
                    else:print(Path+"\\AMDFLASH\\"+Board_version+".bin or _16.bin"+" can't find.")
                else:print(Path+"\\Capsule\\Windows\\"+Board_version+"(00).bin"+" can't find.");\
                    print(Path+"\\Capsule\\Windows\\"+Board_version+"(00).inf"+" can't find.");print(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat"+" can't find.");
            else:print(Path+" can't find.")
    # Check AMD Project S27 S29 T25 T27 T26
    if (Platform_Flag(NewProcPkgInfo) == "S27") or (Platform_Flag(NewProcPkgInfo) == "S29") or (Platform_Flag(NewProcPkgInfo) == "T25") or \
        (Platform_Flag(NewProcPkgInfo) == "T27") or (Platform_Flag(NewProcPkgInfo) == "T26"): # G6/G8 AMD T26 S27
        for NProc in NewProcPkgInfo:
            Board_version=NProc[0]+"_"+NewVersion
            Path=".\\"+("_").join(NProc)
            if os.path.isdir(Path):  # these platform could have 00 in the end of version and the beginning of .cat is uppercase
                if (os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+".inf"))\
                    or(os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+NewBuildID[-2:]+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+NewBuildID[-2:]+".inf")) \
                    or(os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"00.bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"00.inf")) :
                    if (os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+".bin") \
                        or os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+"_32.bin") ) \
                        and \
                        (os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"_"+NewVersion+".cat")\
                        or os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+"00.cat")):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+".bin"):
                            if (os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".bin") \
                                or os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+"00.bin" ))\
                                and \
                                (os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".inf")\
                                or os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+"00.inf")):
                                if os.path.isfile(Path+"\\XML\\"+NProc[0]+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z"):
                                        Check="True"
                                    else:print(Path+"DCI.7z"+" can find.")
                                else:print(Path+"\\XML\\"+NProc[0]+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Board_version+"(00).bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Board_version+"(00).inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Board_version+".bin"+" can't find.")
                    else:print(Path+"\\AMDFLASH\\"+Board_version+".bin or _32.bin"+" can't find.")
                else:print(Path+"\\Capsule\\Windows\\"+Board_version+"(00).bin"+" can't find.");\
                    print(Path+"\\Capsule\\Windows\\"+Board_version+"(00).inf"+" can't find.");print(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+"(0).cat"+" can't find.");
            else:print(Path+" can't find.")
    if Check=="True":   print("New release Pkg made "+Fore.GREEN+"successfully.\n")
    else:   print("New release Pkg made "+Fore.RED+"failed.")


''' 待處理 修改過後的程式碼
def file_exists(path, filenames):
    for filename in filenames:
        if os.path.isfile(os.path.join(path, filename)):
            return True
    return False

def check_files_exist(path, file_groups):
    for files in file_groups:
        if not file_exists(path, files):
            print(f"{path}{' or '.join(files)} can't find.")
            return False
    return True

def CheckPkg(NewProcPkgInfo):
    Check = "False"
    for NProc in NewProcPkgInfo:
        Board_version = NProc[2] + "_" + NProc[3]
        Path = ".\\" + "_".join(NProc)
        if not os.path.isdir(Path):
            print(Path + " can't find.")
            continue

        FPTW_files = [f"{Board_version}_12.bin", f"{Board_version}_16.bin", f"{Board_version}_32.bin"]
        global_bios_files = [f"{Board_version}_16.bin", f"{Board_version}_32.bin"]
        HPFWUPDREC_files = [f"{Board_version}.bin", f"{Board_version}.inf", f"{Board_version}00.bin", f"{Board_version}00.inf"]
        XML_files = [f"{Board_version}.xml"]
        DCI_files = ["DCI.7z", "DCI.zip"]

        if not check_files_exist(Path + "\\FPTW", [FPTW_files]):
            continue
        if not check_files_exist(Path + "\\Global\\BIOS", [global_bios_files]):
            continue
        if not check_files_exist(Path + "\\HPFWUPDREC", [HPFWUPDREC_files]):
            continue
        if not check_files_exist(Path + "\\XML", [XML_files]):
            continue
        if os.path.isfile(Path + "DCI.7z") or os.path.isfile(Path + "DCI.zip"):
            print(Path + "DCI can find.")
            continue

        Check = "True"
        break

    if Check == "True":
        print("New release Pkg made " + Fore.GREEN + "successfully.\n")
    else:
        print("New release Pkg made " + Fore.RED + "failed.")
'''