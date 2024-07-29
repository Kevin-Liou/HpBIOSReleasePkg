# Define a list of all Intel and AMD platforms, just modify this list when new platforms are added.
Intel_Platforms = ["Intel G" + str(i) for i in range(3, 13) if i not in (7, 10, 11)]  # G3 to G12 Remove G7, G10, G11
AMD_Platforms   = ["AMD G" + str(i) for i in range(4, 9) if i not in (7, 9, 10, 11)] # G4 to G8 Remove G7
AMD_ProjectID   = ["Q26", "Q27", "R24", "R26", "S25", "S27", "S29", "T25", "T26", "T27", "X26", "X27"]
# Define a list of all Intel and AMD platforms, just modify this list when new platforms are added.

# Dynamically generate subsets
def get_platforms_subset(all_platforms, start, end=None):
    start_index = all_platforms.index(start)
    end_index = len(all_platforms) if end is None else all_platforms.index(end) + 1
    return set(all_platforms[start_index:end_index])

# Intel specific platform subsets, don't modify.
Intel_Platforms_G3G4    = get_platforms_subset(Intel_Platforms, "Intel G3", "Intel G4")
Intel_Platforms_G4later = get_platforms_subset(Intel_Platforms, "Intel G4")
Intel_Platforms_G5later = get_platforms_subset(Intel_Platforms, "Intel G5")
Intel_Platforms_G6later = get_platforms_subset(Intel_Platforms, "Intel G6")
Intel_Platforms_G9later = get_platforms_subset(Intel_Platforms, "Intel G9")
Intel_Platforms_G12later = get_platforms_subset(Intel_Platforms, "Intel G12")

# AMD specific platform subsets, don't modify.
AMD_Platforms_R24later  = get_platforms_subset(AMD_ProjectID, "R24")
AMD_Platforms_R26later  = get_platforms_subset(AMD_ProjectID, "R26")
AMD_Platforms_ExceptR24 = set(AMD_ProjectID) - {"R24"}


# Check platform is Intel or AMD.
def Platform_Flag(IDCheck):
    if type(IDCheck) == list:
        for ID in IDCheck:
            if "Q26" in ID:
                return "Q26"
            if "Q27" in ID:
                return "Q27"
            if "R26" in ID:
                return "R26"
            if "R24" in ID:
                return "R24"
            if "S25" in ID:
                return "S25"
            if "S27" in ID:
                return "S27"
            if "S29" in ID:
                return "S29"
            if "T25" in ID:
                return "T25"
            if "T26" in ID:
                return "T26"
            if "T27" in ID:
                return "T27"
            if "X26" in ID:
                return "X26"
            if "X27" in ID:
                return "X27"
            if "P10" in ID:
                return "Intel G3"
            if ("Q10" in ID) or ("Q11" in ID) or ("Q21" in ID) or ("Q22" in ID) or ("Q23" in ID) or ("Q35" in ID):
                return "Intel G4"
            if ("R11" in ID) or ("R21" in ID) or ("R22" in ID) or ("R23" in ID):
                return "Intel G5"
            if ("S10" in ID) or ("S11" in ID) or ("S21" in ID) or ("S22" in ID) or ("S23" in ID):
                return "Intel G6"
            if ("T11" in ID) or ("T21" in ID) or ("T22" in ID):
                return "Intel G8"
            if ("U11" in ID) or ("U21" in ID) or ("U22" in ID) or ("U23" in ID):
                return "Intel G9"
            if ("V11" in ID) or ("V21" in ID) or ("V22" in ID) or ("V23" in ID):
                return "Intel G10"
            if ("X11" in ID) or ("X21" in ID) or ("X22" in ID) or ("X23" in ID):
                return "Intel G12"
    else:
        if "Q26" in IDCheck:
            return "Q26"
        if "Q27" in IDCheck:
            return "Q27"
        if "R26" in IDCheck:
            return "R26"
        if "R24" in IDCheck:
            return "R24"
        if "S25" in IDCheck:
            return "S25"
        if "S27" in IDCheck:
            return "S27"
        if "S29" in IDCheck:
            return "S29"
        if "T25" in IDCheck:
            return "T25"
        if "T26" in IDCheck:
            return "T26"
        if "T27" in IDCheck:
            return "T27"
        if "X26" in IDCheck:
            return "X26"
        if "X27" in IDCheck:
            return "X27"
        if "P10" in IDCheck:
            return "Intel G3"
        if ("Q10" in IDCheck) or ("Q11" in IDCheck) or ("Q21" in IDCheck) or ("Q22" in IDCheck) or ("Q23" in IDCheck) or ("Q35" in IDCheck):
            return "Intel G4"
        if ("R11" in IDCheck) or ("R21" in IDCheck) or ("R22" in IDCheck) or ("R23" in IDCheck):
            return "Intel G5"
        if ("S10" in IDCheck) or ("S11" in IDCheck) or ("S21" in IDCheck) or ("S22" in IDCheck) or ("S23" in IDCheck):
            return "Intel G6"
        if ("T11" in IDCheck) or ("T21" in IDCheck) or ("T22" in IDCheck):
            return "Intel G8"
        if ("U11" in IDCheck) or ("U21" in IDCheck) or ("U22" in IDCheck) or ("U23" in IDCheck):
            return "Intel G9"
        if ("V11" in IDCheck) or ("V21" in IDCheck) or ("V22" in IDCheck) or ("V23" in IDCheck):
            return "Intel G10"
        if ("X11" in IDCheck) or ("X21" in IDCheck) or ("X22" in IDCheck) or ("X23" in IDCheck):
            return "Intel G12"

# Debug code.
# if __name__ == '__main__':
#     print("Intel_Platforms: ", Intel_Platforms)
#     print("Intel_Platforms_G3G4: ", Intel_Platforms_G3G4)
#     print("Intel_Platforms_G4later: ", Intel_Platforms_G4later)
#     print("Intel_Platforms_G5later: ", Intel_Platforms_G5later)
#     print("Intel_Platforms_G9later: ", Intel_Platforms_G9later)
#     print("AMD_Platforms: ", AMD_Platforms)
#     print("AMD_Platforms_R24later: ", AMD_Platforms_R24later)
#     print("AMD_Platforms_R26later: ", AMD_Platforms_R26later)
#     print("AMD_Platforms_ExceptR24: ", AMD_Platforms_ExceptR24)