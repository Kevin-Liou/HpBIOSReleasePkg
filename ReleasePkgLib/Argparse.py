import logging
import argparse
from colorama import Fore


def argparse_function(ver):
    parser = argparse.ArgumentParser(
        prog='ReleasePkg.py',
        description='''
        Release Package Creation Tool.

        This script is designed for creating release packages for various platforms.
        Currently supports Intel (G4, G5, G6, G8, G9, G10) and AMD (G4, G5, G6, G8) platforms.
        The script facilitates the process of packaging, renaming, and checking various files and folders
        associated with BIOS release packages.
        ''',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("-d", "--debug", help = 'Show debug message.', action = "store_true")
    parser.add_argument("-v", "--version", action = "version", version = ver)

    args = parser.parse_args()
    #logging.debug(f"args: {args}")
    if args.debug:
        Debug_Format = "%(levelname)s, %(funcName)s: %(message)s"
        logging.basicConfig(level = logging.DEBUG, format = Debug_Format) # Debug use flag
        print(Fore.RED + "Enable debug mode...")
        args = "Debug mode"
    else:
        args = "Release mode"
    return args