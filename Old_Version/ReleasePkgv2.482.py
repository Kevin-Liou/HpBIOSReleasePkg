#Make Release Pkg Script.	v2.482

# coding=UTF-8
from multiprocessing import Pool, Manager
from colorama import init, Fore, Back, Style
import sys, os, re, shutil, zipfile, ftplib, time, multiprocessing
Versionfilelist=["BUFF2.nsh", "Buff2All.nsh", "Update32.bat", "Update64.bat", "UpdateEFI.nsh"]
# In add the Projects to be here, and the following BoradID should also add.
BoardID=["Q10", "Q11", "Q21", "Q22", "Q23"]
CWD=os.getcwd()
Dirlist=os.listdir(CWD)
FTP_host="ftp.usa.hp.com"
FTP_username="sign_ron"
FTP_pwd="5gg9*0UA"

def InputStr(Str):# Input Str, "Str" is message
	if sys.version > '3':
		OutStr=input(Str);
	else:
		OutStr=raw_input(Str);
	return OutStr

def Unzip(Zipname):
	with zipfile.ZipFile(Zipname, 'r') as myzip:
		for file in myzip.namelist():
			myzip.extract(file,File_dir)
		myzip.close()

def ChangebuildID(Proc, filelist):# Modify Update Version
	Filter="_"
	for Filename in filelist:
		File=open(Filename,"r+")
		Strlist=File.read().split(Filter)
		if OldVersion in Strlist:
			for Str in range(len(Strlist)):
				if Strlist[Str]==OldVersion:
					Strlist[Str]=NewVersion
		elif NewVersion in Strlist:
			print("Now Pkg Version is already "+Strlist[1]+".")
			File.close()
			break
		else:
			print("Old Version "+OldVersion+" Can't find, Now Pkg Version is "+Strlist[1]+".")
			File.close()
			sys.exit()
		Strlist=Filter.join(Strlist)
		File.seek(0,0)
		File.write(Strlist)
		File.close()
	print("Modify "+Proc[0]+" Pkg Version (.nsh & .bat) to "+Proc[3]+" succeeded.\n")

def RemoveFileInDir(targetDir, version):# Remove file
	RemoveRule=[".7z", ".xlsm", ".cat", ".cer", ".pfx", ".pvk", "_12.bin", "_16.bin", "_32.bin", version+".bin", version+".inf", version+".xml"]
	for root,dirs,files in os.walk(targetDir):
		for name in files:# Here are the rules for remove
			Path=CWD+"\\"+os.path.join(root, name)
			for Rule in RemoveRule:
				if os.path.isfile(Path) and name.find(Rule)>0:
					os.remove(Path)
					print(os.path.join(root, name)+"  remove succeeded.")

def Copytree(sourcePath, targetPath):# Copy old version folder to new version folder.
	print("Start Copy "+sourcePath.split("\\")[-1]+" to "+targetPath.split("\\")[-1]+", Please wait.....")
	shutil.copytree(sourcePath, targetPath)# Copy to new Pkg
	print("Copy Pkg "+sourcePath.split("\\")[-1]+" to "+targetPath.split("\\")[-1]+" succeeded.\n")

def CopyFiles(sourceFolder,targetFolder):# Cpoy Fv file to Pkg
	sourcefullpath=CWD+"\\"+sourceFolder+"\\"
	targetfullpath=CWD+"\\"+targetFolder+"\\"
	for root,dirs,files in os.walk(sourcefullpath+"\\Combined"):
		for name in files:# Combined copy to Capsule
			shutil.copy(sourcefullpath+"\\Combined\\"+name, targetfullpath+"\\Capsule")
			print(sourceFolder+"\\Combined\\"+name+" to "+targetFolder+"\\Capsule"+" Copy succeeded.")
		for name in files:# Combined copy to HPFWUPDREC
			if name.find(".bin")>0 or name.find(".inf")>0:
				shutil.copy(sourcefullpath+"\\Combined\\"+name, targetfullpath+"\\HPFWUPDREC")
				print(sourceFolder+"\\Combined\\"+name+" to "+targetFolder+"\\HPFWUPDREC"+" Copy succeeded.")
	for name in os.listdir(sourcefullpath):# Bin file copy to FPTW&Global
		if name.find("_12.bin")>0 or name.find("_16.bin")>0 or name.find("_32.bin")>0:
			shutil.copy(sourcefullpath+name, targetfullpath+"\\FPTW")
			print(sourceFolder+"\\"+name+" to "+targetFolder+"\\FPTW"+" Copy succeeded.")
		if Proc[2]=="Q23" and name.find("_16.bin")>0:# If 16MB BIOS Please add it
			shutil.copy(sourcefullpath+name, targetfullpath+"\\Global\\BIOS")
			print(sourceFolder+"\\"+name+" to "+targetFolder+"\\Global\\BIOS"+" Copy succeeded.\n")
		elif name.find("_32.bin")>0:
			shutil.copy(sourcefullpath+name, targetfullpath+"\\Global\\BIOS")
			print(sourceFolder+"\\"+name+" to "+targetFolder+"\\Global\\BIOS"+" Copy succeeded.\n")
		if name.find(".xml")>0:# Copy to XML
			shutil.copy(sourcefullpath+name, targetfullpath+"\\XML")
			print(sourceFolder+"\\"+name+" to "+targetFolder+"\\XML"+" Copy succeeded.")

def FindFvFolder():# Find Fv Folder, Add to Matchlist
	Matchlist=[]
	for Process in Processlist:
		for Dir in Dirlist:
			if NewBuildID=="" or NewBuildID=="0000":
				Match="Fv_"+Process+"_"+NewVersion+"_0000_BuildJob"
			else:
				Match="Fv_"+Process+"_"+NewVersion+"_"+NewBuildID+"_BuildJob"
			if Match in Dir:
				if not str(Dir).find(".zip")>0 and not str(Dir).find(".7z")>0:
					Matchlist.append(Dir)
	return Matchlist

def FindFvZip():# Find Fv Zip file
	Matchlist=[]
	for Process in Processlist:
		for Dir in Dirlist:
			if NewBuildID=="" or NewBuildID=="0000":
				Match="Fv_"+Process+"_"+NewVersion+"_0000_BuildJob"
			else:
				Match="Fv_"+Process+"_"+NewVersion+"_"+NewBuildID+"_BuildJob"
			if Match in Dir:
				if str(Dir).find(".zip")>0 or str(Dir).find(".7z")>0:
					Matchlist.append(Dir)
	return Matchlist

def Ftp_connect(host, username, pwd):# Connect to FTP...
	try:
		Ftp=ftplib.FTP(host)
		Ftp.login(username, pwd)
		return Ftp
	except Exception as err:
		print("\nPlease check the network connection status.\nUnable to download fv Zip file from ftp now.\n")
		sys.exit()

def Ftp_callback(chunk, dst_file, lists, chunklist):
	max_arrow=50  # 設定 '>' 有多少個
	listdate="False"  # 檢查list的狀態
	listnum=0
	tmp=0
	if not hasattr(Ftp_callback, "progress_percent"):  # 初始化
		Ftp_callback.progress_percent=0
	dst_file.write(chunk)  # 把 chunk 寫入到檔案中
	if len(chunklist)==0:  # 假如chunklist中沒有資料，先寫入
		chunklist.append((str(os.getpid()), 0))  # 防止chunk二次增加，先寫入0
	for pid, data in chunklist:  # 尋找目前進程號，chunk疊加在寫回list
		if str(os.getpid())==pid:
			data+=len(chunk)
			listdate="True"
			chunklist[listnum]=(str(os.getpid()), data)
		listnum+=1
	if listdate=="False":  # 假如沒有找到目前進程號，新增資料
		chunklist.append((str(os.getpid()), len(chunk)))
	if len(chunklist)>0:
		for i in range(len(chunklist)):  # 把所有chunk相加
			tmp+=chunklist[i][1]
		lists[1]=tmp
		tmp=lists[1]*100/lists[0]  # 得出目前要顯示的%數
		if Ftp_callback.progress_percent!=tmp:
			Ftp_callback.progress_percent=tmp
			num_arrow=int(Ftp_callback.progress_percent/2)  #要顯示的">"
			num_line=max_arrow-num_arrow  # 要顯示的"-"
			#  要輸出的進度與%數，\r 是回到開頭在顯示，一直循環
			sys.stdout.write("\r["+">"*num_arrow+"-"*num_line+"]"+"%.2f" % Ftp_callback.progress_percent+"%")
			if int(Ftp_callback.progress_percent)==100:# If 100%
				lists[1]=0
				Ftp_callback.progress_percent=0
			sys.stdout.flush()  # 刷新輸出

def Ftp_download(Proc, lists, chunklist):# If Fv folder and .zip can't find, Download from HP FTP
	Ftp=Ftp_connect(FTP_host, FTP_username, FTP_pwd)  # FTP 連線
	Ftpfilelist=[]
	Ftp_dir="/ProductionReleaseDT/"+Proc[2]+"_"+Proc[0]
	Ftp.cwd(Ftp_dir)
	Ftp.dir(Ftpfilelist.append)  # 寫入此資料夾下的所有檔案
	if len(Proc)==4:
		Ftpfilelist=[File for File in Ftpfilelist if "Fv_"+Proc[2]+"_"+Proc[3]+"_0000_BuildJob_" in File]
		if len(Ftpfilelist)==0:  # 如果找不到檔案
			print("Download "+"Fv_"+Proc[2]+"_"+Proc[3]+"_0000_BuildJob"+" Failed, Is not in the ftp site.")
			Ftp.quit()
			return
	if len(Proc)==5:
		Ftpfilelist=[File for File in Ftpfilelist if "Fv_"+Proc[2]+"_"+Proc[3]+"_"+Proc[4]+"_BuildJob_" in File]
		if len(Ftpfilelist)==0:  # 如果找不到檔案
			print("Download "+"Fv_"+Proc[2]+"_"+Proc[3]+"_"+Proc[4]+"_BuildJob"+" Failed, Is not in the ftp site.")
			Ftp.quit()
			return
	Fvfile=Ftpfilelist[0].split("Fv_");
	Fvfile="Fv_"+Fvfile[1]  # 指定Fv檔案的全名
	Ftp.voidcmd("TYPE I")
	File_size=Ftp.size(Ftp_dir+"/"+Fvfile)  # 算出指定Fv檔案的大小
	if len(lists)==0:
		lists.append(File_size)
		lists.append(0)
	else:
		lists[0]+=File_size  # 計算出大小總和
	if not os.path.isfile(CWD+"\\"+Fvfile):  # 檢查Fv檔案是不是已經存在
		print("Start Download "+Fvfile+"\tfrom FTP, Please wait.....")
		f=open(Fvfile, "wb")  # 開啟Fv檔案準備寫入
		try:  # 開始下載&呼叫Ftp_callback載入進度
			time.sleep(0.5)
			Ftp.retrbinary("RETR %s/%s" % (Ftp_dir, Fvfile), lambda chunk:Ftp_callback(chunk, f, lists, chunklist))
		except ftplib.error_perm:  # 如果出現錯誤
			print("Download "+Fvfile+" Failed, Is not in the ftp site.")
			f.close()
			Ftp.quit()
			return
		else:
			print()
			print("Download %s\tsuccessed\n" % Fvfile)
		f.close()  # 關閉FTP
		lists.append(Fvfile)  # 把下載完成的檔案加入list
	else:	print(Fvfile+" have already exists.")  # 如果Fv檔案已經存在
	Ftp.quit()

def Ftp_multi(NewProcSplit):
	lists=Manager().list()  # 定義一個進程池可以共享的lists
	chunklist=Manager().list()
	po=Pool(len(BoardID))  # 定義一個進程池，最大進程數
	for Proc in NewProcSplit:
		# Pool.apply_async(要調用的目標,(傳遞給目標的參數元祖,))
		# 每次循環將會用空閒出來的子進程去調用目標
		po.apply_async(Ftp_download, (Proc, lists, chunklist))
	print("----start----")
	po.close()  # 關閉進程池，關閉後po不再接收新的請求
	po.join()  # 等待po中所有子進程執行完成，必須放在close語句之後
	print("-----end-----")
	return lists  # 返回進程池共用的lists

def CheckPkg():# Check New Release Pkg is OK?
	Check="False"
	for Proc in NewProcSplit:
		Boardversion=Proc[2]+"_"+Proc[3]
		Path=CWD+"\\"+("_").join(Proc)
		if os.path.isdir(Path):
			if os.path.isfile(Path+"\\Capsule\\"+Boardversion+".bin") and os.path.isfile(Path+"\\Capsule\\"+Boardversion+".inf") and os.path.isfile(Path+"\\Capsule\\"+Proc[2].lower()+"_"+NewVersion+".cat"):
				if os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_12.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_32.bin"):
					if os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_32.bin"):
						if os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"):
							if os.path.isfile(Path+"\\XML\\"+Boardversion+".xml"):
								if not os.path.isfile(Path+"DCI.7z"):
									Check="True"
								else:print(Path+"DCI.7z"+" can find.")
							else:print(Path+"\\XML\\"+Boardversion+".xml"+" can't find.")
						else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.");
					else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
				else:print(Path+"\\FPTW\\"+Boardversion+"_12.bin"+" can't find.")
			else:print(Path+"\\Capsule\\"+Boardversion+".bin"+" can't find.");print(Path+"\\Capsule\\"+Boardversion+".inf"+" can't find.");print(Path+"\\Capsule\\"+Proc[2].lower()+"_"+NewVersion+".cat"+" can't find.");
		else:print(Path+" can't find.")
	if Check=="True":	print("New release Pkg made "+Fore.GREEN+"successfully.")
	else:	print("New release Pkg made "+Fore.RED+"failed.")

	# Script Start
if __name__ == '__main__':
	init(autoreset=True)
	multiprocessing.freeze_support()# For windows do multiprocessing.
	print("==============Make Release Pkg Script   v2.482==========================================")
	print("==============Input Information=========================================================")
	OldVersion=InputStr("OldVersion:")# Input OldVersion & NewVersion & NewBuildID
	NewVersion=InputStr("NewVersion:")
	NewBuildID=InputStr("NewBuildID:")# ex: 020106_"0001"
	Process=InputStr("                                   (Can Multiple Select)\
		\nPlease Enter Projects To Processed"+str(BoardID)+":")# Input need Process boardID, Can multiple choice
	
	Processlist=Process.upper().split()
	if OldVersion=="":print("\nPlease Input OldVersion.");sys.exit()
	elif NewVersion=="":print("\nPlease Select NewVersion.");sys.exit()

	print("\n==============Find Need Process Old Pkg=================================================")
	OldBuildID="0";ProcName=[];ProcSplit=[]
	print("Start Find Process Old Pkg.")
	for Process in Processlist:
		temp=[];
		for Dir in Dirlist:
			if not Dir.split("_")[0]=="Fv" and not Dir.find(".7z")>0 and not Dir.find(".zip")>0:
				if Process+"_"+OldVersion in Dir:
					temp.append(Dir)
		if len(temp)>1 and OldBuildID=="0":
			OldBuildID=InputStr("\n"+str(temp)+"\n"+Process+"_"+OldVersion+" Please Select OldBuildID:")
			if OldBuildID=="" or OldBuildID=="0000":
				temp=[a for a in temp if len(a.split("_"))==4]
			else:
				temp=[a for a in temp if Process+"_"+OldVersion+"_"+OldBuildID in a]
		if len(temp)>1 and OldBuildID!="0":
			if OldBuildID=="" or OldBuildID=="0000":
				temp=[a for a in temp if len(a.split("_"))==4]
			else:
				temp=[a for a in temp if Process+"_"+OldVersion+"_"+OldBuildID in a]
			if not len(temp)==1:
				temp=[Dir for Dir in Dirlist if Process+"_"+OldVersion in Dir]
				OldBuildID=InputStr("\n"+str(temp)+"\n"+Process+"_"+OldVersion+" Please Select OldBuildID:")
				if OldBuildID=="" or OldBuildID=="0000":
					temp=[a for a in temp if len(a.split("_"))==4]
				else:
					temp=[a for a in temp if Process+"_"+OldVersion+"_"+OldBuildID in a]
		if len(temp)==0:
			print(Process+"_"+OldVersion+" Old Pkg folder can't find, Please check.")
		else:
			ProcName.append(temp[0])
	print("\nYour need process old Pkg:\n"+str(ProcName))

	ProcSplit=[Proc.split("_") for Proc in ProcName]# Split to process
	NewProcSplit=ProcSplit[:]# New Pkg Name List
	if NewBuildID=="" or NewBuildID=="0000":
		for Proc in range(len(NewProcSplit)):
			NewProcSplit[Proc]=NewProcSplit[Proc][:3]
			NewProcSplit[Proc].append(NewVersion)
	else:
		for Proc in range(len(NewProcSplit)):
			NewProcSplit[Proc]=NewProcSplit[Proc][:3]
			NewProcSplit[Proc].append(NewVersion)
			NewProcSplit[Proc].append(NewBuildID)

	print("\n==============Find Fv Folder Or Zip File================================================")
	Matchfolderlist=FindFvFolder()
	Matchziplist=FindFvZip()

	print("Your Fv Folder: ", end=" ");print(Matchfolderlist)
	print("Your Fv Zip File: ", end=" ");print(str(Matchziplist)+"\n")
	# If can't find Fv folder or Zip file.
	if len(Matchfolderlist)==0 and len(Matchziplist)==0:
		print("Can't find Fv folder and zip file.\nDownload Fv files from Production Release FTP.\n")
		temp=Ftp_multi(NewProcSplit)[:]
		for name in temp:
			if str(name).find(".zip")>0:
				Matchziplist.append(name)
	# Number of Fv folders not match Projectlist
	elif len(Matchfolderlist)<len(Processlist) and len(Matchfolderlist)!=0 and len(Matchziplist)<len(Processlist):
		print("Number of Fv folders not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
		temp=Ftp_multi(NewProcSplit)[:]
		for name in temp:
			if str(name).find(".zip")>0:
				Matchziplist.append(name)
	# Can't find Fv folders and Number of Fv Zip files not match Projectlist
	elif len(Matchfolderlist)==0 and len(Matchziplist)<len(Processlist):
		print("Fv Zip files not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
		temp=Ftp_multi(NewProcSplit)[:]
		for name in temp:
			if str(name).find(".zip")>0:
				Matchziplist.append(name)
	# If Fv folder already exists
	elif len(Matchfolderlist)==len(Processlist):
		print("Fv folder already exists.\n")
	# Find Fv Zip file Start extracting

	if len(Matchziplist)>0:
		print("\nFind Fv Zip File, Start Extracting.")
		for Zipname in Matchziplist:
			File_dir=Zipname.split(".")[0]
			if not os.path.isdir(CWD+"\\"+File_dir):
				Unzip(Zipname)
				shutil.move(CWD+"\\"+Zipname, CWD+"\\"+File_dir)
				print(Zipname+" Extract succeeded.")
				Matchfolderlist.append(File_dir)
			else:
				print("Fv folder "+File_dir+" already exists, Remove Fv zip file.")
				os.remove(CWD+"\\"+Zipname)
		print("\nNow Your Fv Folder: ", end=" ");print(Matchfolderlist)
	else:	print("\nNow Your Fv Folder: ", end=" ");print(Matchfolderlist)

	print("\n==============Find New Pkg Or Add New Pkg===============================================")
	for Proc in range(len(ProcSplit)):# How much Old Version folder
		OldVersionPath=CWD+"\\"+("_").join(ProcSplit[Proc])
		NewVersionPath=CWD+"\\"+("_").join(NewProcSplit[Proc])
		if not os.path.isdir(NewVersionPath):# Check NewVersion Folder Exist
			if not os.path.isdir(OldVersionPath+"\\"+("_").join(ProcSplit[Proc])):# Check Old Pkg is in folder???
				Copytree(OldVersionPath, NewVersionPath)
			elif not os.path.isdir(OldVersionPath+"\\FPTW"):
				if os.path.isdir(OldVersionPath+"\\"+("_").join(ProcSplit[Proc])):# Check Old Pkg is in folder???
					Copytree(OldVersionPath+"\\"+("_").join(ProcSplit[Proc]), NewVersionPath)
				else:	print("Pkg "+("_").join(ProcSplit[Proc])+" can't find.")
			else:	print("Pkg "+("_").join(ProcSplit[Proc])+" can't find.")
		else:	print("Pkg "+NewVersionPath.split("\\")[-1]+" already exists.")
	if len(Matchfolderlist)==0:	print("Can't find anything Fv folder.\n")

	print("\n==============Modify Pkg Update Version=================================================")
	for Proc in NewProcSplit:# Pkg Modify Update Version
		Path=CWD+"\\"+("_").join(Proc)
		if os.path.isdir(Path+"\\FPTW"):# Check Folder Exist
			os.chdir(Path)
			ReleaseNote=[ReleaseNote for ReleaseNote in os.listdir(Path) if "release note.docx" in ReleaseNote]
			if len(ReleaseNote)==1:
				os.rename(ReleaseNote[0], ("_").join(Proc)+" release note.docx")
			print("ReleaseNote Rename to "+("_").join(Proc)+" release note.docx"+" succeeded.")
			os.chdir(Path+"\\FPTW")
			ChangebuildID(Proc, Versionfilelist)
		else:	print("Pkg "+("_").join(Proc)+" can't find.\n")
	os.chdir(CWD)

	print("\n==============Remove Pkg Old File=======================================================")
	for Proc in NewProcSplit:
		targetfolder=("_").join(Proc)
		if (os.path.isdir(targetfolder)):
			RemoveFileInDir(targetfolder, OldVersion)
			print()
		else:	print("Pkg "+("_").join(Proc)+" can't find.\n")
	os.chdir(CWD)

	print("==============Fv File Rename And Copy To Pkg============================================")
	for Fv in Matchfolderlist:
		for Proc in NewProcSplit:
			if Fv.split("_")[1]==Proc[2]:
				if os.path.isdir(CWD+"\\"+Fv):
					Path=CWD+"\\"+Fv
					os.chdir(Path)
					Boardversion=Proc[2]+"_"+Proc[3]
					if "Fv_Q23_"+NewVersion+"_" in Fv:# For 16MB BIOS
					# If add 16MB BIOS Please add [or "Fv_QXX_"+NewVersion+"_" in Matchlist[Match]] in it
						if os.path.isfile(Path+"\\"+Boardversion+"_12.bin") and os.path.isfile(Path+"\\"+Boardversion+".xml"):# If Alreadly renamed
							print(Boardversion+"_12.bin & "+Boardversion+".xml alreadly renamed.")
						if os.path.isfile(Path+"\\"+Boardversion+".bin") and os.path.isfile(Path+"\\"+Proc[2]+".xml"):
							os.rename(Path+"\\"+Boardversion+".bin", Path+"\\"+Boardversion+"_12.bin")# Rename Fv folder 2 files
							os.rename(Path+"\\"+Proc[2]+".xml", Path+"\\"+Boardversion+".xml")
							print(Boardversion+"_12.bin & "+Boardversion+".xml rename succeeded.")
						if os.path.isfile(Path+"\\"+Boardversion+".inf"):
							if os.path.isdir(CWD+"\\"+("_").join(Proc)):# Check Pkg Folder Exist
								CopyFiles(Fv, ("_").join(Proc))
							else:	print("Pkg "+("_").join(Proc)+" can't find.")
					else:
						if os.path.isfile(Path+"\\"+Boardversion+"_16.bin") and os.path.isfile(Path+"\\"+Boardversion+".xml"):# If Alreadly renamed
							print(Boardversion+"_16.bin & "+Boardversion+".xml alreadly renamed.")
						if os.path.isfile(Path+"\\"+Boardversion+".bin") and os.path.isfile(Path+"\\"+Proc[2]+".xml"):
							os.rename(Path+"\\"+Boardversion+".bin", Path+"\\"+Boardversion+"_16.bin")# Rename Fv folder 2 files
							os.rename(Path+"\\"+Proc[2]+".xml", Path+"\\"+Boardversion+".xml")
							print(Boardversion+"_16.bin & "+Boardversion+".xml rename succeeded.")
						if os.path.isfile(Path+"\\"+Boardversion+".inf"):
							if os.path.isdir(CWD+"\\"+("_").join(Proc)):# Check Pkg Folder Exist
								CopyFiles(Fv, ("_").join(Proc))
							else:	print("Pkg "+("_").join(Proc)+" can't find.")
				else:	print("Need to be processed Fv folder:"+Fv+" can't find.\n")
	if len(Matchfolderlist)==0 or len(Processlist)==0:	print("Can't find anything Fv folder.\n")

	print("==============End=======================================================================")
	CheckPkg()# Check new release Pkg is OK?
	print("\nFinally pkg please compare with leading project.\n")
	os.system('Pause')

	sys.exit()
	# Script End