import logging
import argparse
from colorama import Fore, Style, init

# Initialize colorama, automatically redefine color
init(autoreset=True)

# Custom Formatter, set different colors according to log level
class ColorFormatter(logging.Formatter):
    FORMATS = {
        logging.DEBUG: Fore.CYAN + "%(levelname)s, %(funcName)s: %(message)s" + Style.RESET_ALL,
        logging.INFO: Fore.GREEN + "%(levelname)s, %(funcName)s: %(message)s" + Style.RESET_ALL,
        logging.WARNING: Fore.YELLOW + "%(levelname)s, %(funcName)s: %(message)s" + Style.RESET_ALL,
        logging.ERROR: Fore.RED + "%(levelname)s, %(funcName)s: %(message)s" + Style.RESET_ALL,
        logging.CRITICAL: Fore.MAGENTA + "%(levelname)s, %(funcName)s: %(message)s" + Style.RESET_ALL,
    }

    def format(self, record):
        log_fmt = self.FORMATS.get(record.levelno, self._fmt)
        formatter = logging.Formatter(log_fmt)
        return formatter.format(record)

def argparse_function(ver):
    parser = argparse.ArgumentParser(
        prog='ReleasePkg.py',
        description='''\
Release Package Creation Tool.

This script is designed for creating release packages for various platforms.
Currently supports Intel (G4, G5, G6, G8, G9, G10) and AMD (G4, G5, G6, G8) platforms.
The script facilitates the process of packaging, renaming, and checking various files and folders
associated with BIOS release packages.
        ''',
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument("-d", "--debug", help='Show debug message.', action="store_true")
    parser.add_argument("-v", "--version", action="version", version=ver)

    args = parser.parse_args()

    # Setting up logging handler and using custom formatter
    handler = logging.StreamHandler()
    handler.setFormatter(ColorFormatter())
    logger = logging.getLogger()  # Get root logger
    logger.handlers = []  # Remove the default handler
    logger.addHandler(handler)

    if args.debug:
        logger.setLevel(logging.DEBUG)
        logger.debug(f"args: {args}")
        print(Fore.RED + "Enable debug mode...")
        args = "Debug mode"
    else:
        args = "Release mode"
    return args