import sys, os, logging
from ftplib import FTP, error_perm
from time import sleep
from multiprocessing import Pool, Manager

from ReleasePkgLib import *


# Connect to FTP...
def Ftp_connect(Server_info):
    try:
        logging.debug('Try to connect Ftp...')
        Ftp = FTP(Server_info["host"])
        Ftp.login(Server_info["username"], Server_info["pwd"])
        print("Connect to FTP......succeeded.")
        return Ftp
    except Exception as err:
        ExitProgram("\nPlease check the network connection status.\nUnable to download fv Zip file from ftp now.\n")


def Ftp_get_file_list(line):
    file_arr = Ftp_get_filename(line)
    if file_arr[1] not in ['.', '..']:
        file_list.append(file_arr)


def Ftp_get_filename(line):
    pos = line.rfind(':')
    while line[pos] != ' ':
        pos += 1
    while line[pos] == ' ':
        pos += 1
    file_arr = [line[0], line[pos:]]
    return file_arr


# FTP callback 進度條.
def Ftp_callback(chunk, dst_file, lists, chunk_list):
    max_arrow = 50  # 設定 '>' 有多少個
    list_date = "False"  # 檢查list的狀態
    list_num = 0
    tmp = 0
    if not hasattr(Ftp_callback, "progress_percent"):  # 初始化
        Ftp_callback.progress_percent = 0
    dst_file.write(chunk)  # 把 chunk 寫入到檔案中
    if len(chunk_list) == 0:  # 假如chunk_list中沒有資料，先寫入
        chunk_list.append((str(os.getpid()), 0))  # 防止chunk二次增加，先寫入0
    for pid, data in chunk_list:  # 尋找目前進程號，chunk疊加在寫回list
        if str(os.getpid()) == pid:
            data += len(chunk)
            list_date = "True"
            chunk_list[list_num] = (str(os.getpid()), data)
        list_num += 1
    if list_date == "False":  # 假如沒有找到目前進程號，新增資料
        chunk_list.append((str(os.getpid()), len(chunk)))
    if len(chunk_list) != -1:
        for i in range(len(chunk_list)):  # 把所有chunk相加
            tmp += chunk_list[i][1]
        lists[1] = tmp
        tmp = lists[1] * 100 / lists[0]  # 得出目前要顯示的%數
        if Ftp_callback.progress_percent != tmp:
            Ftp_callback.progress_percent = tmp
            num_arrow = int(Ftp_callback.progress_percent / 2)  # 要顯示的">"
            num_line = max_arrow - num_arrow  # 要顯示的"-"
            #  要輸出的進度與%數，\r 是回到開頭在顯示，一直循環
            sys.stdout.write(
                "\r[" + ">" * num_arrow + "-" * num_line + "]" + "%.2f" % Ftp_callback.progress_percent + "%")
            if int(Ftp_callback.progress_percent) == 100:  # If 100%
                lists[1] = 0
                Ftp_callback.progress_percent = 0
            sys.stdout.flush()  # 刷新輸出


def Ftp_download_Test(local_dir, remote_dir, Test_Server_info):
    global file_list
    logging.debug('Ftp_connect...')
    Ftp = Ftp_connect(Test_Server_info)  # FTP 連線
    try:
        Ftp.cwd(remote_dir)
    except:
        print('Directory %s not exist,Continue...' % remote_dir)
        return
    if not os.path.isdir(local_dir):
        os.makedirs(local_dir)
    print('Switch to directory %s' % Ftp.pwd())
    file_list = []
    Ftp.dir(Ftp_get_file_list)
    remote_names = file_list
    for item in remote_names:
        filetype = item[0]
        filename = item[1]
        local = os.path.join(local_dir, filename)
        if filetype == 'd':
            Ftp_download_Test(local, remote_dir + '/' + filename, Test_Server_info)
        elif filetype == '-':
            print('>>>>>>>>>>>>Download file %s ... ...' % local)
            file_handler = open(local, 'wb')
            Ftp.retrbinary('RETR %s' % (filename), file_handler.write)
            file_handler.close()
    Ftp.cwd('..')
    print('Return to upper directory %s' % Ftp.pwd())


# If Fv folder and .zip can't find, Download from HP FTP.
def Ftp_download(NProc, lists, chunk_list, Server_info):
    logging.debug('Ftp_connect...')
    Ftp = Ftp_connect(Server_info)  # FTP 連線
    Ftp_file_list = []
    Ftp_dir = "/ProductionReleaseDT/" + NProc[2] + "_" + NProc[0]
    if Server_info["type"] == "Test":
        Ftp_dir = "/Test Release/" + NProc[2]
    Ftp.cwd(Ftp_dir)
    Ftp.dir(Ftp_file_list.append)  # 寫入此資料夾下的所有檔案
    if len(NProc) == 4:
        Ftp_file_list = [File for File in Ftp_file_list if "Fv_" + NProc[2] + "_" + NProc[3] + "_0000_BuildJob_" in File]
        if len(Ftp_file_list) == 0:  # 如果找不到檔案
            print("Download " + "Fv_" + NProc[2] + "_" + NProc[3] + "_0000_BuildJob" + " Failed, Is not in the ftp site.")
            Ftp.quit()
            return False
    if len(NProc) == 5:
        Ftp_file_list = [File for File in Ftp_file_list if "Fv_" + NProc[2] + "_" + NProc[3] + "_" + NProc[4] + "_BuildJob_" in File]
        if len(Ftp_file_list) == 0:  # 如果找不到檔案
            print("Download " + "Fv_" + NProc[2] + "_" + NProc[3] + "_" + NProc[4] + "_BuildJob" + " Failed, Is not in the ftp site.")
            Ftp.quit()
            return False
    Fv_file = Ftp_file_list[0].split("Fv_")
    Fv_file = "Fv_" + Fv_file[1]  # 指定Fv檔案的全名
    Ftp.voidcmd("TYPE I")
    File_size = Ftp.size(Ftp_dir + "/" + Fv_file)  # 算出指定Fv檔案的大小
    if len(lists) == 0:
        lists.append(File_size)
        lists.append(0)
    else:
        lists[0] += File_size  # 計算出大小總和
    if not os.path.isfile(".\\" + Fv_file):  # 檢查Fv檔案是不是已經存在
        print("Start Download " + Fv_file + "\tfrom FTP, Please wait.....")
        f = open(Fv_file, "wb")  # 開啟Fv檔案準備寫入
        try:  # 開始下載&呼叫Ftp_callback載入進度
            sleep(0.5)
            Ftp.retrbinary("RETR %s/%s" % (Ftp_dir, Fv_file), lambda chunk: Ftp_callback(chunk, f, lists, chunk_list))
        except error_perm:  # 如果出現錯誤
            print("Download " + Fv_file + " Failed, Is not in the ftp site.")
            f.close()
            Ftp.quit()
            return
        else:
            print()
            print("Download %s\tsuccessed\n" % Fv_file)
        f.close()  # 關閉FTP
        lists.append(Fv_file)  # 把下載完成的檔案加入list
    else:
        print(Fv_file + " have already exists.")  # 如果Fv檔案已經存在
    Ftp.quit()


# FTP 多進程下載
def Ftp_multi(NewProcPkgInfo, Production_Server_info, Test_Server_info):
    lists = Manager().list()  # 定義一個進程池可以共享的lists
    chunk_list = Manager().list()
    po = Pool(10)  # 定義一個進程池，最大進程數
    Choose = InputStr("Download from ProductionSign key:1, TestSign key:2 :")
    if Choose == "1":
        logging.debug('Download from ProductionSign start.')
        Server_info = Production_Server_info
        for NProc in NewProcPkgInfo:
            # Pool.apply_async(要調用的目標,(傳遞給目標的參數元祖,))
            # 每次循環將會用空閒出來的子進程去調用目標
            po.apply_async(Ftp_download, (NProc, lists, chunk_list, Server_info))
    else:
        logging.debug('Download from TestSign start.')
        Server_info = Test_Server_info
        for NProc in NewProcPkgInfo:
            po.apply_async(Ftp_download, (NProc, lists, chunk_list, Server_info))
            ''' 因為test sign ftp也有放zip檔後就不用下載資料夾
            if len(NProc) == 4:
                local_dir = '.' + r'\Fv_' + NProc[2] + '_' + NProc[3] + '_0000'  # 本地目錄
                remote_dir = '/Test Release/' + NProc[2] + '/' + NProc[3][0: 2] + "." + NProc[3][2:4] + "." + NProc[3][4:6] + '_0000'  # 遠程目錄
            else:
                local_dir = '.' + r'\Fv_' + NProc[2] + '_' + NProc[3] + '_' + NProc[4]  # 本地目錄
                remote_dir = '/Test Release/' + NProc[2] + '/' + NProc[3][0: 2] + "." + NProc[3][2:4] + "." + NProc[3][4:6] + '_' + NProc[4]  # 遠程目錄
            po.apply_async(Ftp_download_Test, (local_dir, remote_dir, Test_Server_info))
            '''
    print("----start----")
    po.close()  # 關閉進程池，關閉後po不再接收新的請求
    po.join()  # 等待po中所有子進程執行完成，必須放在close語句之後
    print("-----end-----")
    return lists  # 返回進程池共用的lists