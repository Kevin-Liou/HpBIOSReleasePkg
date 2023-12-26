import argparse
import logging
from colorama import Fore


def argparse_function(ver):
    parser = argparse.ArgumentParser(prog='ReleasePkg.py', description='Tutorial')
    parser.add_argument("-d", "--debug", help='Show debug message.', action="store_true")
    parser.add_argument("-v", "--version", action="version", version=ver)
    args = parser.parse_args()
    if args.debug:
        Debug_Format = "%(levelname)s, %(funcName)s: %(message)s"
        logging.basicConfig(level=logging.DEBUG, format=Debug_Format) # Debug use flag
        print(Fore.RED + "Enable debug mode...")
        return "Debug mode"
    if args.version:
        return ver