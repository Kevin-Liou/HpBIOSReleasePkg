import os
from colorama import Fore

from ReleasePkgLib import *
from .Platform import Platform_Flag


# Check New Release Pkg is OK?
def CheckPkg(NewProcPkgInfo):
    Check="False"
    for NProc in NewProcPkgInfo:
        Board_version=NProc[2]+"_"+NProc[3]
        Path=".\\"+("_").join(NProc)
        if ((Platform_Flag(NewProcPkgInfo) == "Intel G3") or (Platform_Flag(NewProcPkgInfo) == "Intel G4")):
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

        if ((Platform_Flag(NewProcPkgInfo) == "Intel G5") or (Platform_Flag(NewProcPkgInfo) == "Intel G6") or (Platform_Flag(NewProcPkgInfo) == "Intel G8") or (Platform_Flag(NewProcPkgInfo) == "Intel G9")or \
            (Platform_Flag(NewProcPkgInfo) == "Intel G10")):
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
    if (Platform_Flag(NewProcPkgInfo) == "R24") or (Platform_Flag(NewProcPkgInfo) == "Q26") or (Platform_Flag(NewProcPkgInfo) == "Q27"):
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
    if (Platform_Flag(NewProcPkgInfo) == "R26") or (Platform_Flag(NewProcPkgInfo) == "S25"):# G5 and G6 AMD R26 AMD S25
        for NProc in NewProcPkgInfo:
            Board_version=NProc[0]+"_"+NewVersion
            Path=".\\"+("_").join(NProc)
            if os.path.isdir(Path):
                if (os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+".inf"))\
                or(os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+NewBuildID[-2:]+".bin") and os.path.isfile(Path+"\\Capsule\\Windows\\"+Board_version+NewBuildID[-2:]+".inf")) :
                    if (os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+".bin") or os.path.isfile(Path+"\\AMDFLASH\\"+Board_version+"_16.bin")) and \
                        os.path.isfile(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat"):
                        if os.path.isfile(Path+"\\Global\\BIOS\\"+Board_version+".bin"):
                            if os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Board_version+".inf"):
                                if os.path.isfile(Path+"\\XML\\"+NProc[0]+".xml"):
                                    if not os.path.isfile(Path+"DCI.7z"):
                                        Check="True"
                                    else:print(Path+"DCI.7z"+" can find.")
                                else:print(Path+"\\XML\\"+NProc[0]+".xml"+" can't find.")
                            else:print(Path+"\\HPFWUPDREC\\"+Board_version+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Board_version+".inf"+" can't find.");
                        else:print(Path+"\\Global\\BIOS\\"+Board_version+".bin"+" can't find.")
                    else:print(Path+"\\AMDFLASH\\"+Board_version+".bin or _16.bin"+" can't find.")
                else:print(Path+"\\Capsule\\Windows\\"+Board_version+".bin"+" can't find.");\
                    print(Path+"\\Capsule\\Windows\\"+Board_version+".inf"+" can't find.");print(Path+"\\Capsule\\Windows\\"+NProc[0].lower()+"_"+NewVersion+".cat"+" can't find.");
            else:print(Path+" can't find.")
    if (Platform_Flag(NewProcPkgInfo) == "S27") or (Platform_Flag(NewProcPkgInfo) == "S29") or (Platform_Flag(NewProcPkgInfo) == "T25") or \
        (Platform_Flag(NewProcPkgInfo) == "T27") or (Platform_Flag(NewProcPkgInfo) == "T26"):#G6/G8 AMD T26 S27
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