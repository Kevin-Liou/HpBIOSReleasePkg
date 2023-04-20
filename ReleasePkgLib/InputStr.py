import sys


# Input Str, "Str" is message.
def InputStr(Str):
    if sys.version > '3':
        OutStr = input(Str)
    else:
        OutStr = raw_input(Str)
    return OutStr