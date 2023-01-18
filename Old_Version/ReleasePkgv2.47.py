#Make Release Pkg script.	v2.47

# coding=UTF-8
from multiprocessing import Pool, Manager
import sys, os, re, shutil, zipfile, ftplib, time
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
	print("Modify "+Proc[0]+" Pkg Version (.nsh & .bat) to "+Proc[3]+" succeeded.")

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
			if Match in Dir and not Dir.find(".zip")>0:
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
			if Match in Dir and Dir.find(".zip")>0:
				Matchlist.append(Dir)
	return Matchlist

def Ftpconnect(host, username, pwd):# Connect to FTP...
	try:
		Ftp=ftplib.FTP(host)
		Ftp.login(username, pwd)
		return Ftp
	except Exception as err:
		print("\nPlease check the network connection status.\nUnable to download fv Zip file from ftp now.\n")
		sys.exit()

def Ftp_callback(chunk, file_size, dst_file, lists):
	max_arrow=50
	if not hasattr(Ftp_callback, "progress_size"):
		Ftp_callback.progress_size=0
	if not hasattr(Ftp_callback, "progress_percent"):
		Ftp_callback.progress_percent=0
	dst_file.write(chunk)
	if len(lists)==1:
		time.sleep(1)
		lists.append(len(chunk))
	else:
		lists[1]+=len(chunk)# chunk is size of data received each time
	tmp=lists[1]*100/lists[0]
	if Ftp_callback.progress_percent!=tmp:
		Ftp_callback.progress_percent=tmp
		num_arrow=int(Ftp_callback.progress_percent/2)# ">" numbers
		num_line=max_arrow-num_arrow# "-" numbers
		sys.stdout.write("\r["+">"*num_arrow+"-"*num_line+"]"+"%.2f" % Ftp_callback.progress_percent+"%")# show [>>>]xx%, \r is back to start.
		if int(Ftp_callback.progress_percent)==100:# If 100%
			Ftp_callback.progress_size=0
			Ftp_callback.progress_percent=0
		sys.stdout.flush()

def Ftp_download(Proc, lists):# If Fv folder and .zip can't find, Download from HP FTP
	global Matchziplist
	Ftp=Ftpconnect(FTP_host, FTP_username, FTP_pwd)
	Ftpfilelist=[]
	Ftp_dir="/ProductionReleaseDT/"+Proc[2]+"_"+Proc[0]
	Ftp.cwd(Ftp_dir)
	Ftp.dir(Ftpfilelist.append)# Write FTP dir to Ftpfilelist
	if len(Proc)==4:
		Ftpfilelist=[File for File in Ftpfilelist if "Fv_"+Proc[2]+"_"+Proc[3]+"_0000_BuildJob_" in File]
		if len(Ftpfilelist)==0:
			print("Download "+"Fv_"+Proc[2]+"_"+Proc[3]+"_0000_BuildJob"+" Failed, Is not in the ftp site.")
			Ftp.quit()
			return
	else:
		Ftpfilelist=[File for File in Ftpfilelist if "Fv_"+Proc[2]+"_"+Proc[3]+"_"+Proc[4]+"_BuildJob_" in File]
		if len(Ftpfilelist)==0:
			print("Download "+"Fv_"+Proc[2]+"_"+Proc[3]+"_"+Proc[4]+"_BuildJob"+" Failed, Is not in the ftp site.")
			Ftp.quit()
			return
	Data=Ftpfilelist[0].split("Fv_");
	Data="Fv_"+Data[1]# Set Fv file full name
	Ftp.voidcmd("TYPE I")
	File_size=Ftp.size(Ftp_dir+"/"+Data)
	if len(lists)==0:
		lists.append(File_size)
	else:
		lists[0]+=File_size
	if not os.path.isfile(CWD+"\\"+Data):# Check Fv zip file exists.
		print("Start Download "+Data+" from FTP, Please wait.....")
		f=open(Data, "wb")
		try:# Start download Fv file, and callback chunk size
			Ftp.retrbinary("RETR %s/%s" % (Ftp_dir, Data), lambda chunk:Ftp_callback(chunk, File_size, f, lists))
		except ftplib.error_perm:
			print("Download "+Data+" Failed, Is not in the ftp site.")
			f.close()
			Ftp.quit()
			return
		else:
			print()
			print("Download %s successed\n" % Data)
		f.close()
		lists.append(Data)
	else:	print(Data+" have already exists.")
	Ftp.quit()

def Ftp_multi(NewProcSplit):
	lists=Manager().list() # 定義一個進程池可以共享的list
	po=Pool(3)  # 定義一個進程池，最大進程數3
	for Proc in NewProcSplit:
		# Pool.apply_async(要調用的目標,(傳遞給目標的參數元祖,))
		# 每次循環將會用空閒出來的子進程去調用目標
		po.apply_async(Ftp_download, (Proc, lists))
	print("----start----")
	po.close()  # 關閉進程池，關閉後po不再接收新的請求
	po.join()  # 等待po中所有子進程執行完成，必須放在close語句之後
	print("-----end-----")
	return lists[2:]

def CheckPkg():# Check New Release Pkg is OK?
	Check="false"
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
									Check="true"
								else:print(Path+"DCI.7z"+" can find.")
							else:print(Path+"\\XML\\"+Boardversion+".xml"+" can't find.")
						else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.\n");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.\n");
					else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
				else:print(Path+"\\FPTW\\"+Boardversion+"_12.bin"+" can't find.")
			else:print(Path+"\\Capsule\\"+Boardversion+".bin"+" can't find.\n");print(Path+"\\Capsule\\"+Boardversion+".inf"+" can't find.\n");print(Path+"\\Capsule\\"+Proc[2].lower()+"_"+NewVersion+".cat"+" can't find.\n");
		else:print(Path+" can't find.\n")
	if Check=="true":	print("New release Pkg made successfully.")
	else:	print("New release Pkg made failed.")

# Script Start
if __name__ == '__main__':
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
		temp=[];temp=[Dir for Dir in Dirlist if Process+"_"+OldVersion in Dir]
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
			NewProcSplit[Proc].append(NewVersion+"_"+NewBuildID)
	
	print("\n==============Find Fv Folder Or Zip File================================================")
	Matchfolderlist=FindFvFolder()
	Matchziplist=FindFvZip()

	print("Your Fv Folder: ", end=" ");print(Matchfolderlist)
	print("Your Fv Zip File: ", end=" ");print(str(Matchziplist)+"\n")
	# If can't find Fv folder or Zip file.
	if len(Matchfolderlist)==0 and len(Matchziplist)==0:
		print("Can't find Fv folder and zip file.\nDownload Fv files from Production Release FTP.\n")
		Matchziplist=Ftp_multi(NewProcSplit)[2:]
	# Number of Fv folders not match Projectlist
	elif len(Matchfolderlist)<len(Processlist) and len(Matchfolderlist)!=0 and len(Matchziplist)<len(Processlist):
		print("Number of Fv folders not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
		Matchziplist=Ftp_multi(NewProcSplit)[2:]
	# Can't find Fv folders and Number of Fv Zip files not match Projectlist
	elif len(Matchfolderlist)==0 and len(Matchziplist)<len(Processlist):
		print("Fv Zip files not match Projectlist.\nDownload Fv files from Production Release FTP.\n")
		Matchziplist=Ftp_multi(NewProcSplit)[2:]
	# If Fv folder already exists
	elif len(Matchfolderlist)==len(Processlist):
		print("Fv folder already exists.\n")
	# Find Fv Zip file Start extracting

	if len(Matchziplist)>0:
		print("Find Fv Zip File, Start extracting.")
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
			Copytree(OldVersionPath, NewVersionPath)
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
	print("\nFinally pkg please compare with leading project.")
	os.system('Pause')

	sys.exit()
	# Script End
