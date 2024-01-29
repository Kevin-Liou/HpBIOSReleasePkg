import logging
import os
import struct
from colorama import Fore


from ReleasePkgLib import *
from .Platform import *


def GetBinaryData(Match_folder_list, FindData, Offset, DataSize, Unpack_format):
    logging.debug("Get Binary Data Start.")
    BinaryData = ""  # Initialize BinaryData outside the loop
    for Fv in Match_folder_list:
        Path = ".\\" + Fv
        if os.path.isdir(Path):
            files = os.listdir(Path)
            for name in files:
                if "_32.bin" in name:
                    logging.debug("find BIOS binary: " + name)
                    with open(os.path.join(Path, name), "rb") as BinaryFile:
                        FileData = BinaryFile.read()
                        Data_Str_index = FileData.find(FindData)
                        if Data_Str_index == -1:
                            logging.debug('Can not find Data in Binary')
                        else:
                            Data_End_index = Data_Str_index + len(FindData) + Offset + DataSize
                            VersionData = FileData[Data_Str_index + len(FindData) + Offset:Data_End_index]
                            byte_sequence = struct.unpack(Unpack_format, VersionData)
                            BinaryData = '.'.join(map(str, byte_sequence))
    if BinaryData == "":
        logging.debug("Get Binary Data failed.")
    return BinaryData



def GetMrcVersion(Match_folder_list):
    logging.debug("Get Mrc Version Start.")
    Version = ""
    for Fv in Match_folder_list:
        if Platform_Flag(Fv) == "Intel G5": #Block Intel G5 for MRC Version error.
            return (Version)
        if Platform_Flag(Fv) == "Intel G6": #Intel G6 offset different with other generations.
            Offset = 4
        else:
            Offset = 0
    FindData = b'MRCVER_'
    UnpackDataSize = '>4B'
    DataSize = 0x04
    Version = GetBinaryData(Match_folder_list, FindData, Offset, DataSize, UnpackDataSize)
    if Version == "":
        logging.debug("Get MRC Version fail.")
    return (Version)


def GetMEVersion(Match_folder_list):
    logging.debug("Get ME Version Start.")
    Version = ""
    Offset = 0x94
    FindData = b'RBEP.man'
    UnpackDataSize = '<4H'
    DataSize = 8
    Version = GetBinaryData(Match_folder_list, FindData, Offset, DataSize, UnpackDataSize)
    if Version == "":
        logging.debug("Get ISH Version fail.")
    return (Version)


def GetIshVersion(Match_folder_list):
    logging.debug("Get ISH Version Start.")
    Version = ""
    Offset = 0x94
    for Fv in Match_folder_list:
        if Platform_Flag(Fv) == "Intel G8": #Modify Intel G8 ISH offset.
            Offset = 0x64
    FindData = b'ISHC.man'
    UnpackDataSize = '<4H'
    DataSize = 8
    Version = GetBinaryData(Match_folder_list, FindData, Offset, DataSize, UnpackDataSize)
    if Version == "":
        logging.debug("Get ISH Version fail.")
    return (Version)


def GetPmcVersion(Match_folder_list):
    logging.debug("Get PMC Version Start.")
    Version = ""
    Offset = 0x94
    for Fv in Match_folder_list:
        if Platform_Flag(Fv) == "Intel G6": #Modify Intel G6 PMC offset.
            Offset = 0x64
        elif Platform_Flag(Fv) == "Intel G8": #Modify Intel G8 PMC offset.
            Offset = 0xF4
    FindData = b'PMCP.man'
    UnpackDataSize = '<4H'
    DataSize = 8
    Version = GetBinaryData(Match_folder_list, FindData, Offset, DataSize, UnpackDataSize)
    if Version == "":
        logging.debug("Get PMC Version fail.")
    else :
        version_list = Version.split('.')
        new_version_list = ['{:02d}'.format(int(x)) for x in version_list]
        Version = '.'.join(new_version_list)
    return (Version)


def GetNphyVersion(Match_folder_list):
    logging.debug("Get ME Nphy Version Start.")
    Version = ""
    Offset = 0xC4
    FindData = b'NPHY.man'
    UnpackDataSize = '<4H'
    DataSize = 8
    Version = GetBinaryData(Match_folder_list, FindData, Offset, DataSize, UnpackDataSize)
    if Version == "":
        logging.debug("Get NPHY Version fail.")
    return (Version)


def GetSphyVersion(Match_folder_list):
    logging.debug("Get ME Sphy Version Start.")
    Version = ""
    Offset = 0x184
    FindData = b'SPHY.man'
    UnpackDataSize = '<4H'
    DataSize = 8
    Version = GetBinaryData(Match_folder_list, FindData, Offset, DataSize, UnpackDataSize)
    if Version == "":
        logging.debug("Get SPHY Version fail.")
    else :
        version_list = Version.split('.')
        new_version_list = ['{:02d}'.format(int(x)) for x in version_list]
        Version = '.'.join(new_version_list)
    return (Version)


def GetPchcVersion(Match_folder_list):
    logging.debug("Get ME PCHC Version Start.")
    Version = ""
    Offset = 0x94
    FindData = b'PCHC.man'
    UnpackDataSize = '<4H'
    DataSize = 8
    Version = GetBinaryData(Match_folder_list, FindData, Offset, DataSize, UnpackDataSize)
    if Version == "":
        logging.debug("Get PCHC Version fail.")
    else :
        version_list = Version.split('.')
        new_version_list = ['{:02d}'.format(int(x)) for x in version_list]
        Version = '.'.join(new_version_list)
    return (Version)