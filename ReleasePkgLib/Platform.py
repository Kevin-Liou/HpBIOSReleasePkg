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