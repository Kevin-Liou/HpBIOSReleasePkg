# ReleasePkgLib/__init__.py

from multiprocessing import freeze_support
from colorama import init


from .Lib import MatchMultipleFolder, ChangeBuildID, RemoveOldFileInDir, Copy_Release_Folder, Copy_Release_Files, Copy_Release_Files_AMD, FindFvFolder, FindFvZip
from .InputStr import InputStr
from .Zip import UnZip, PrintZipAllInfo, GetZipDateInfo
from .Ftp import Ftp_multi, Ftp_download, Ftp_download_Test, Ftp_connect, Ftp_get_file_list, Ftp_get_filename, Ftp_callback
from .Excel import CheckBiosBuildDate, ModifyExcelData, FindOldMEVersion, GetBinaryData, GetMrcVersion, GetMEVersion, GetIshVersion, GetPmcVersion, GetNphyVersion, GetSphyVersion, GetPchcVersion, PrintBiosBuildDate, CheckFileChecksum, PrintBiosBinaryChecksum, CheckMEVersion, ModifyReleaseNote
from .Platform import Platform_Flag
from .Argparse import argparse_function
from .Version import Version
from .Check import CheckPkg, CheckPkg_AMD
from .Tool_Version_Compare import GetFileInfo, ReadToolVersionTable, SetToolVersionTable, CompareInfo, ChangeVersionInfo, ChangeDataInfo


def main_init():
    init(autoreset=True)# colorama
    freeze_support()# For windows do multiprocessing.
    args = argparse_function(Version())