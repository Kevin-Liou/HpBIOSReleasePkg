# ReleasePkgLib/__init__.py

from multiprocessing import freeze_support
from colorama import init

# Import libraries and functions
from .Lib import (
    ExitProgram,
    MatchMultipleFolder,
    CheckBiosVersion,
    SafeRename,
    ChangeBuildID,
    RemoveOldFileInDir,
    Copy_Release_Folder,
    Copy_Release_Files,
    Copy_Release_Files_AMD,
    FindFvFolder,
    FindFvZip,
)

from .Zip import (
    UnZip,
    PrintZipAllInfo,
    GetZipDateInfo,
)

from .Ftp import (
    Ftp_multi,
    Ftp_download,
    Ftp_download_Test,
    Ftp_connect,
    Ftp_get_file_list,
    Ftp_get_filename,
    Ftp_callback,
)

from .Excel import (
    CheckBiosBuildDate,
    ModifyExcelData,
    FindOldMEVersion,
    PrintBiosBuildDate,
    CheckFileChecksum,
    PrintBiosBinaryChecksum,
    CheckMEVersion,
    ModifyReleaseNote,
)

from .Platform import (
    Platform_Flag,
    Intel_Platforms,
    Intel_Platforms_G3G4,
    Intel_Platforms_G4later,
    Intel_Platforms_G5later,
    Intel_Platforms_G6later,
    Intel_Platforms_G9later,
    AMD_Platforms,
    AMD_Platforms_R24later,
    AMD_Platforms_R26later,
    AMD_Platforms_G12later,
    AMD_Platforms_ExceptR24,
)

from .ToolVersionCompare import (
    GetFileInfo,
    ReadToolVersionTable,
    SetToolVersionTable,
    CompareInfo,
    ChangeVersionInfo,
    ChangeDataInfo,
)

from .GetBinFWVersion import (
    GetBinaryData,
    GetMrcVersion,
    GetMEVersion,
    GetIshVersion,
    GetPmcVersion,
    GetNphyVersion,
    GetSphyVersion,
    GetPchcVersion,
)

from .Config import Config_init, Config_debug, config_data
from .Argparse import argparse_function
from .Check import CheckPkg, CheckPkg_AMD
from .Version import Version
from .InputStr import InputStr

def Main_init():
    # Initialize colorama and multiprocessing
    init(autoreset=True)  # colorama
    freeze_support()  # For windows multiprocessing

    # Parse arguments and initialize configuration
    args = argparse_function(Version())
    Config_data = Config_init()

    return args, Config_data
