import os
import sys
import json

# Defining Configuration Data
config_data = {
    # Supported BIOS Board ID.
    "BoardID": ["P10", "Q10~11", "Q21~23", "Q26~27", "R11", "R21~24", "R26",
                "S10~11", "S21~23", "S25~S29", "T11,T21~22", "T25~27",
                "U11,U21~23", "V11, V21~V23"
    ],
    # AMD platform ID and platform name.
    "AMDProjectName": {
        "Q26": "ScottyRr",
        "Q27": "Scotty",
        "R24": "Worf",
        "R26": "Riker",
        "S25": "DoppioPco",
        "S27": "DoppioRn",
        "S29": "CubanoRn",
        "T25": "DoppioCzn",
        "T26": "CubanoCzn",
        "T27": "DoppioR8"
    },
    # The content requires modifications to the bios version files.
    "VersionFileList": [
        "BUFF2.nsh", "Buff2All.nsh", "Update32.bat", "Update64.bat",
        "UpdateEFI.nsh", "Update32_vPro.bat", "Update64_vPro.bat", "UpdateEFI_vPro.nsh"
    ],
    # Keywords for files you don't want to delete, Priority over "Remove_file_rule".
    "NotRemoveFileRule": [
        "Note", "note", "History", "How to Flash", "AMT_CFG",
        "logo", "sign.bin", "HPSignME", "Batch"
    ],
    # Keywords for files you want to delete, can use regular expressions.
    "RemoveFileRule": [
        "DCI..+", ".cer", ".pfx", ".pvk", ".xlsm", ".log", "Pvt.bin","metainfo.xml", "Build.Log", "TBT_RTD3.+", "QA's report", "QAreport",
        "\\d\\d_\\d\\d_\\d\\d.bin", "\\d\\d_\\d\\d_\\d\\d.cat", "\\d\\d_\\d\\d_\\d\\d.inf", # U21_02100400.bin
        "\\d{6}.bin", "\\d{6}.cat", "\\d{6}.cab", "\\d{6}.inf",
        "\\d{4}_12.bin", "\\d{4}_16.bin", "\\d{4}_32.bin", # U21_021004_32.bin
        "\\w\\d{2}_\\d{4}.bin", "\\w\\d{2}_\\d{4}.cat", "\\w\\d{2}_\\d{4}.xml", "\\w\\d{2}_\\d{4}.inf", "\\w\\d{2}_\\d{6}.xml",
        "P00\\w{3}-\\w{3}.zip", "HP_\\w+_\\w+_\\w+_\\w+_\\w+_\\w+_\\d+.+",
        "ME_+[0-9]+[.]+[0-9]+[.]+[0-9]+[.]+[0-9]+.bin"
    ],
    # Production release package ftp server.
    "ProductionReleaseServer": {
        "type": "Production",
        "host": "ftp.usa.hp.com",
        "username": "sign_ron",
        "password": "7gg9*0UA"
    },
    # Test release package ftp server.
    "TestReleaseServer": {
        "type": "Test",
        "host": "ftp.usa.hp.com",
        "username": "bios15ws",
        "password": "e.QV9ra}"
    }
}


# Check and create Config.json file.
def Config_init():
    config_path = "Config.json"

    # 從 config_data 字典中提取鍵作為必需的項目集合
    required_keys = set(config_data.keys())

    # 檢查 Config.json 文件是否存在
    if os.path.isfile(config_path):
        with open(config_path, 'r') as file:
            Config_data = json.load(file)

        # 檢查所有必要的項目是否存在
        if not all(key in Config_data for key in required_keys):
            # 如果有遺漏的項目，刪除舊的 Config.json 並創建新的
            os.remove(config_path)
            with open(config_path, 'w') as f:
                f.write(json.dumps(config_data, indent=4))
            print("Config file already exists, But missing items found. Config file has been recreated.")
            return config_data
        else:
            print("Config file already exists, And all required items are present in the Config file.")
            return Config_data

    else:
        # 如果 Config.json 文件不存在，創建它
        with open(config_path, 'w') as f:
            f.write(json.dumps(config_data, indent=4))
        print("Config file has been created.")
        return config_data


# Check and create Config_debug.json file.
def Config_debug():
    config_file = 'Config_debug.json'

    # 檢查 Config_debug.json 是否存在
    if not os.path.exists(config_file):
        # 配置文件不存在，提示用戶輸入訊息
        OldVersion = input("(Debug)Enter OldVersion: ")
        NewVersion = input("(Debug)Enter NewVersion: ")
        NewBuildID = input("(Debug)Enter NewBuildID: ")
        ProcessProject = input("(Debug)Enter ProcessProject: ")
        ProcessProjectList = ProcessProject.upper().split() # ex:['U21', 'U23']
        if ProcessProject == "":
            input("\n(Debug)Please Input Project.(ex. U21)")
            sys.exit()

        # 保存輸入的訊息到 JSON 文件
        config_data = {
            "OldVersion": OldVersion,
            "NewVersion": NewVersion,
            "NewBuildID": NewBuildID,
            "ProcessProject": ProcessProject,
            "ProcessProjectList": ProcessProjectList
        }

        with open(config_file, 'w') as file:
            json.dump(config_data, file, indent=4)

    else:
        # 讀取現有的配置文件
        with open(config_file, 'r') as file:
            config_data = json.load(file)
            OldVersion = config_data["OldVersion"]
            NewVersion = config_data["NewVersion"]
            NewBuildID = config_data["NewBuildID"]
            ProcessProject = config_data["ProcessProject"]
            ProcessProjectList = config_data["ProcessProjectList"]
        print()
        print("(Debug)OldVersion: " + OldVersion)
        print("(Debug)NewVersion: " + NewVersion)
        print("(Debug)NewBuildID: " + NewBuildID)
        print("(Debug)ProcessProject: " + ProcessProject)
        print("(Debug)ProcessProjectList: " + str(ProcessProjectList))
        print()
        # 用戶確認配置文件中的訊息是否正確
        if input("(Debug)Are you sure these messages are correct? (Y/N): ").upper() == "N":
            # 如果不正確，提示用戶重新輸入
            OldVersion = input("(Debug)Enter OldVersion: ")
            NewVersion = input("(Debug)Enter NewVersion: ")
            NewBuildID = input("(Debug)Enter NewBuildID: ")
            ProcessProject = input("(Debug)Enter ProcessProject: ")
            ProcessProjectList = ProcessProject.upper().split()
            if ProcessProject == "":
                input("\n(Debug)Please Input Project.(ex. U21)")
                sys.exit()

            # 保存輸入的訊息到 JSON 文件
            config_data = {
                "OldVersion": OldVersion,
                "NewVersion": NewVersion,
                "NewBuildID": NewBuildID,
                "ProcessProject": ProcessProject,
                "ProcessProjectList": ProcessProjectList
            }

            with open(config_file, 'w') as file:
                json.dump(config_data, file, indent=4)

    return OldVersion, NewVersion, NewBuildID, ProcessProject, ProcessProjectList