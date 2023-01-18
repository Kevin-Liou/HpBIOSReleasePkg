#Make Release Pkg script.	v2.4

# coding=UTF-8
from __future__ import print_function
import sys, os, re, shutil, zipfile, ftplib, time
Filter="_"
Versionfilelist=["BUFF2.nsh", "Buff2All.nsh", "Update32.bat", "Update64.bat", "UpdateEFI.nsh"]
# In add the Projects to be here, and the following BoradID should also add.
Projectlist=["Oboe", "Sax", "Harp", "Lute", "Mandolin"]
BoardID=["Q10", "Q11", "Q21", "Q22", "Q23"]
Phase="MV"
Phaselise=["DB", "SI", "SI2", "PV", "MV"]
CWD=os.getcwd()
Dirlist=os.listdir(CWD)

try:
	ProductionReleaseFtp=ftplib.FTP("ftp.usa.hp.com")
	ProductionReleaseFtp.login("sign_ron", "5gg9*0UA")
	net="true"
except Exception as err:
	print("\nPlease check the network connection status.\nUnable to download fv Zip file from ftp now.\n");net="false"

def ChangebuildID():# Modify Update Version
	for Filename in Versionfilelist:
		File=open(Filename,"r+")
		FileStr=File.read()
		Strlist=FileStr.split(Filter)
		if (OldVersion in Strlist):
			for a in range(len(Strlist)):
				if Strlist[a]==OldVersion:
					Strlist[a]=NewVersion
			Strlist=Filter.join(Strlist)
			File.seek(0,0)
			File.write(Strlist)
			File.close()
		else:
			print("Old Version "+OldVersion+" Can't find, Now Pkg Version is "+Strlist[1]+".")
			File.close()
			sys.exit()
	if NewBuildID=="" or NewBuildID=="0000":
		print("Modify "+Projectlist[Project]+" Pkg Version (.nsh & .bat) to "+NewVersion+" succeeded.")
	else:
		print("Modify "+Projectlist[Project]+" Pkg Version (.nsh & .bat) to "+NewVersion+"_"+NewBuildID+" succeeded.")

def RemoveFileInDir(targetDir, targetFile):# Remove file
	for root,dirs,files in os.walk(targetDir):
		for name in files:# Here are the rules for remove
			if os.path.isfile(name) or name.find(".7z")>0 or name.find(".xlsm")>0 or name.find(".cat")>0 or name.find(".cer")>0 or name.find(".pfx")>0:
				files=CWD+"\\"+os.path.join(root, name)
				os.remove(files)
				print(os.path.join(root, name)+"  remove succeeded.")
			elif name.find(targetFile+".bin")>0 or name.find("_12.bin")>0 or name.find("_16.bin")>0 or name.find("_32.bin")>0 or name.find(targetFile+".inf")>0 or name.find(targetFile+".xml")>0:
				files=CWD+"\\"+os.path.join(root, name)
				os.remove(files)
				print(os.path.join(root, name)+"  remove succeeded.")
			elif os.path.isfile(CWD+"\\"+targetDir+"\\Capsule\\fwu.pvk") and name.find(".pvk")>0:
				files=CWD+"\\"+targetDir+"\\Capsule\\fwu.pvk"
				os.remove(files)
				print(os.path.join(root, name)+"  remove succeeded.")

def Copytree(sourcePath, targetPath):# Copy old version folder to new version folder.
	print("Start Copy "+sourcePath.split("\\")[-1]+" to "+targetPath.split("\\")[-1]+", Please wait.....")
	shutil.copytree(sourcePath, targetPath)# Copy to new Pkg
	print("Copy Pkg "+sourcePath.split("\\")[-1]+" to "+targetPath.split("\\")[-1]+" succeeded.")

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
		if BoardID[Boardnum]=="Q23" and name.find("_16.bin")>0:# If 16MB BIOS Please add it
			shutil.copy(sourcefullpath+name, targetfullpath+"\\Global\\BIOS")
			print(sourceFolder+"\\"+name+" to "+targetFolder+"\\Global\\BIOS"+" Copy succeeded.\n")
		elif name.find("_32.bin")>0:
			shutil.copy(sourcefullpath+name, targetfullpath+"\\Global\\BIOS")
			print(sourceFolder+"\\"+name+" to "+targetFolder+"\\Global\\BIOS"+" Copy succeeded.\n")
		if name.find(".xml")>0:# Copy to XML
			shutil.copy(sourcefullpath+name, targetfullpath+"\\XML")
			print(sourceFolder+"\\"+name+" to "+targetFolder+"\\XML"+" Copy succeeded.")

def FindFvfolder():# Find Fv Folder, Add to Matchlist
	for Dir in range(len(Dirlist)):
		for Boardnum in range(len(BoardID)):# For all file compare all Borad
			if NewBuildID=="" or NewBuildID=="0000":
				Matchstr="Fv_"+BoardID[Boardnum]+"_"+NewVersion+"_0000_BuildJob"
			else:
				Matchstr="Fv_"+BoardID[Boardnum]+"_"+NewVersion+"_"+NewBuildID+"_BuildJob"
			if Matchstr in Dirlist[Dir]:
				if not Dirlist[Dir].find(".zip")>0:# Find folder not .zip
					if not Dirlist[Dir] in Matchlist:
						Matchlist.append(Dirlist[Dir])
				else:
					if not Dirlist[Dir] in Matchziplist:
						Matchziplist.append(Dirlist[Dir])

def Ftpdownload():# If Fv folder and .zip can't find, Download from HP FTP
	Ftpfilelsit=[]
	if net=="true":
		for num in range(len(Projectlist)):
			if BoardID[num] in Processlist:# Find process BoardID
				ProductionReleaseFtp.cwd("/ProductionReleaseDT/"+BoardID[num]+"_"+Projectlist[num])
				ProductionReleaseFtp.dir(Ftpfilelsit.append)# Write FTP dir to Ftpfilelsit
				if NewBuildID=="" or NewBuildID=="0000":
					Ftpfilelsit=[Filename for Filename in Ftpfilelsit if "Fv_"+BoardID[num]+"_"+NewVersion+"_0000_BuildJob_" in Filename]
					if len(Ftpfilelsit)==0:
						print("Download "+"Fv_"+BoardID[num]+"_"+NewVersion+"_0000_BuildJob"+" Failed, Is not in the ftp site.")
						continue
				else:
					Ftpfilelsit=[Filename for Filename in Ftpfilelsit if "Fv_"+BoardID[num]+"_"+NewVersion+"_"+NewBuildID+"_BuildJob_" in Filename]
					if len(Ftpfilelsit)==0:
						print("Download "+"Fv_"+BoardID[num]+"_"+NewVersion+"_"+NewBuildID+"_BuildJob"+" Failed, Is not in the ftp site.")
						continue
				Data=Ftpfilelsit[0].split("Fv_");Data="Fv_"+Data[1]# Set Fv file full name
				if not os.path.isfile(CWD+"\\"+Data):# Check Fv zip exists
					if len(Compartlist)==0:# If Fv folder not anything exists
						print("Start Download "+Data+" from FTP, Please wait.....")
						f=open(Data, "wb")
						try:ProductionReleaseFtp.retrbinary('RETR ' + Data, f.write)# Start download Fv file
						except ftplib.error_perm:
							print("Download "+Data+" Failed, Is not in the ftp site.")
						print("Download "+Data+" succeeded.")
						f.close()
						Matchziplist.append(Data)
					else:
						if not Data.split(".")[0] in Compartlist:# If Fv folder not match Projectlist
							print("Start Download "+Data+" from FTP, Please wait.....")
							f=open(Data, "wb")
							try:ProductionReleaseFtp.retrbinary("RETR " + Data, f.write)# Start download Fv file
							except ftplib.error_perm:
								print("Download "+Data+" Failed, Is not in the ftp site.")
							print("Download "+Data+" succeeded.")
							f.close()
							Matchziplist.append(Data)
				else: print(Data+" have already exists.")
		ProductionReleaseFtp.close()
	else:	print("Network connection status error.\n")

def CheckPkg():# Check new release Pkg is OK?
	Check="false"
	for ID in range(len(BoardID)):
		if BoardID[ID] in Processlist:
			Boardversion=BoardID[ID]+"_"+NewVersion
			if NewBuildID=="" or NewBuildID=="0000":	Path=CWD+"\\"+Projectlist[ID]+"_"+Phase+"_"+BoardID[ID]+"_"+NewVersion
			else:	Path=CWD+"\\"+Projectlist[ID]+"_"+Phase+"_"+BoardID[ID]+"_"+NewVersion+"_"+NewBuildID
			if os.path.isdir(Path):
				if os.path.isfile(Path+"\\Capsule\\"+Boardversion+".bin") and os.path.isfile(Path+"\\Capsule\\"+Boardversion+".inf") and os.path.isfile(Path+"\\Capsule\\"+BoardID[ID].lower()+"_"+NewVersion+".cat"):
					if os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_12.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\FPTW\\"+Boardversion+"_32.bin"):
						if os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_16.bin") or os.path.isfile(Path+"\\Global\\BIOS\\"+Boardversion+"_32.bin"):
							if os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".bin") and os.path.isfile(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"):
								if os.path.isfile(Path+"\\XML\\"+Boardversion+".xml"):
									Check="true"
								else:print(Path+"\\XML\\"+Boardversion+".xml"+" can't find.")
							else:print(Path+"\\HPFWUPDREC\\"+Boardversion+".bin"+" can't find.\n");print(Path+"\\HPFWUPDREC\\"+Boardversion+".inf"+" can't find.\n");
						else:print(Path+"\\Global\\BIOS\\"+Boardversion+".bin"+" can't find.")
					else:print(Path+"\\FPTW\\"+Boardversion+"_12.bin"+" can't find.")
				else:print(Path+"\\Capsule\\"+Boardversion+".bin"+" can't find.\n");print(Path+"\\Capsule\\"+Boardversion+".inf"+" can't find.\n");print(Path+"\\Capsule\\"+BoardID[ID].lower()+"_"+NewVersion+".cat"+" can't find.\n");
			else:print(Path+" can't find.\n")
	if Check=="true":	print("New release Pkg made successfully.")
	else:	print("New release Pkg made failed.")

# Script Start
if sys.version > '3':OldVersion=input("OldVersion:");NewVersion=input("NewVersion:");NewBuildID=input("NewBuildID:");# ex: 020106_0001
else:	OldVersion=raw_input("OldVersion:");NewVersion=raw_input("NewVersion:");NewBuildID=raw_input("NewBuildID:");

if OldVersion=="":print("\nPlease input OldVersion.");sys.exit()
elif NewVersion=="":print("\nPlease input NewVersion.");sys.exit()

print("\n==============Find Fv Folder or Zip file================================================")
Matchlist=[];Matchziplist=[];Compartlist=[];Processlist=[];
FindFvfolder()
print("Your Fv Folder: ", end=" ");print(Matchlist)
print("Your Fv Zip File: ", end=" ");print(Matchziplist)
# If can't find Fv folder or Zip file.
if len(Matchlist)==0 and len(Matchziplist)==0:
	print("Can't find Fv folder and zip file.\n")
	if sys.version > '3':
		Processlist=input("Download Fv files from Production Release FTP.\nPlease enter projects to processed"+str(BoardID)+":")
		print()
	else:
		Processlist=raw_input("Download Fv files from Production Release FTP.\nPlease enter projects to processed"+str(BoardID)+":")
		print()
	Processlist=Processlist.upper()
	Processlist=Processlist.split()
	Ftpdownload()
# Number of Fv folders not match Projectlist
if len(Matchlist)<len(BoardID) and len(Matchlist)!=0 and len(Matchziplist)<len(BoardID):
	Compartlist=Matchlist
	print("Number of Fv folders not match Projectlist.\n")
	if sys.version > '3':
		Processlist=input("Download Fv files from Production Release FTP.\nPlease enter projects to processed"+str(BoardID)+":")
		print()
	else:
		Processlist=raw_input("Download Fv files from Production Release FTP.\nPlease enter projects to processed"+str(BoardID)+":")
		print()
	Processlist=Processlist.upper()
	Processlist=Processlist.split()
	Ftpdownload()
# Can't find Fv folders and Number of Fv Zip files not match Projectlist
if len(Matchlist)==0 and len(Processlist)==0 and len(Matchziplist)>0:
	Compartlist=Matchlist
	print("Fv Zip files not match Projectlist.\n")
	if sys.version > '3':
		Processlist=input("Download Fv files from Production Release FTP.\nPlease enter projects to processed"+str(BoardID)+":")
		print()
	else:
		Processlist=raw_input("Download Fv files from Production Release FTP.\nPlease enter projects to processed"+str(BoardID)+":")
		print()
	Processlist=Processlist.upper()
	Processlist=Processlist.split()
	Ftpdownload()
# Can't find Fv folders and Number of Fv Zip files not match Projectlist
if len(Matchlist)==0 and len(Matchziplist)==len(BoardID):
	print("You have all projects Fv Zip files, Congratulations.\n")
# Find Fv Zip file Start extracting
FindFvfolder()
if len(Matchziplist)>0:
	for zip in Matchziplist:
		if not zip.split("_")[1] in Processlist:
			Processlist.append(zip.split("_")[1])
	print("Find Fv Zip File, Start extracting.")
	for Zipname in Matchziplist:
		File_dir=Zipname.split(".")[0]
		if not (os.path.isdir(CWD+"\\"+File_dir)):
			with zipfile.ZipFile(Zipname, 'r') as myzip:
				for file in myzip.namelist():	myzip.extract(file,File_dir)
				myzip.close()
			shutil.move(CWD+"\\"+Zipname, CWD+"\\"+File_dir)
			print(Zipname+" Extract succeeded.")
			Matchlist.append(File_dir)
		else:
			print("Fv folder "+File_dir+" already exists, Remove Fv zip file.")
			os.remove(CWD+"\\"+Zipname)
	print("Now Your Fv Folder: ", end=" ");print(Matchlist)
else:	print("Now Your Fv Folder: ", end=" ");print(Matchlist)

print("\n==============Find new Pkg or Add new Pkg===============================================")
OldBuildID=0
for Fvfolder in Matchlist:# How much Fv folder
	for Project in range(len(Projectlist)):
		if BoardID[Project] in Processlist:# Find process BoardID
			if BoardID[Project]==Fvfolder.split("_")[1]:# If Fv folder boardID match
				Check=""
				for list in Phaselise:# Check phase and oldBuild is exist?
					for x in range(0, 10):
						if os.path.isdir(CWD+"\\"+Projectlist[Project]+"_"+list+"_"+BoardID[Project]+"_"+OldVersion):# No build ID 
							Phase=list;Check="NonBuildID"
						elif os.path.isdir(CWD+"\\"+Projectlist[Project]+"_"+list+"_"+BoardID[Project]+"_"+OldVersion+"_000"+str(x)):# On build ID ,Up to 9
							Phase=list;OldBuildID="000"+str(x);Check="BuildID"
				if Check=="NonBuildID":# If BuildID not exists
					OldVersionPkg=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+OldVersion
					OldVersionPath=CWD+"\\"+OldVersionPkg
				elif Check=="BuildID":# If BuildID exists
					OldVersionPkg=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+OldVersion+"_"+OldBuildID
					OldVersionPath=CWD+"\\"+OldVersionPkg
				else:# Can't find anything oldversion folder
					print(Projectlist[Project]+" OldVersion:"+OldVersion+" & BuildID:"+"0000~0009"+" Pkg folder can't find, Please check.")
					print("Skip the "+Projectlist[Project]+" pkg script to continue.")
					continue
				if NewBuildID=="" or NewBuildID=="0000":
					NewVersionPkg=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion
					NewVersionPath=CWD+"\\"+NewVersionPkg
				else:
					NewVersionPkg=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion+"_"+NewBuildID
					NewVersionPath=CWD+"\\"+NewVersionPkg
				if not os.path.isdir(NewVersionPath):# Check NewVersion Folder Exist
					Copytree(OldVersionPath, NewVersionPath)
				else:	print("Pkg "+NewVersionPkg+" already exists.")
if len(Matchlist)==0 or len(Processlist)==0:	print("Can't find anything Fv folder.")

print("\n==============Modify Pkg Update Version=================================================")
for Project in range(len(Projectlist)):# Pkg Modify Update Version
	if BoardID[Project] in Processlist:# Find process BoardID
		Pkgname=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion
		Path=CWD+"\\"+Pkgname
		if NewBuildID=="" or NewBuildID=="0000":
			if (os.path.isdir(Path+"\\FPTW")):# Check Folder Exist
				os.chdir(Path+"\\FPTW")
				ChangebuildID()
				os.chdir(Path)# Start modify .docx file name
				ReleaseNote=[ReleaseNote for ReleaseNote in os.listdir(Path) if "release note.docx" in ReleaseNote][0]
				os.rename(ReleaseNote, Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion+" release note.docx")
				print("ReleaseNote Rename to "+Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion+" release note.docx"+" succeeded.")
			else:	print("Pkg "+Pkgname+" can't find.")
		else:
			if (os.path.isdir(Path+"_"+NewBuildID+"\\FPTW")):# Check Folder Exist
				os.chdir(Path+"_"+NewBuildID+"\\FPTW")
				ChangebuildID()
				os.chdir(Path+"_"+NewBuildID)# Start modify .docx file name
				ReleaseNote=[ReleaseNote for ReleaseNote in os.listdir(Path+"_"+NewBuildID) if "release note.docx" in ReleaseNote][0]
				os.rename(ReleaseNote, Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion+"_"+NewBuildID+" release note.docx")
				print("ReleaseNote Rename to "+Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion+"_"+NewBuildID+" release note.docx"+" succeeded.")
			else:	print("Pkg "+Pkgname+"_"+NewBuildID+" can't find.")
		os.chdir(CWD)

print("\n==============Remove Pkg Old File=======================================================")
for Project in range(len(Projectlist)):# Remove Pkg Old File
	if BoardID[Project] in Processlist:# Find process BoardID
		if NewBuildID=="" or NewBuildID=="0000":
			sourcefolder=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+OldVersion
			targetfolder=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion
		else:
			sourcefolder=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+OldVersion
			targetfolder=Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion+"_"+NewBuildID
		if (os.path.isdir(targetfolder)):
			RemoveFileInDir(targetfolder, OldVersion)
			print()
		else:	print("Pkg "+Projectlist[Project]+"_"+Phase+"_"+BoardID[Project]+"_"+NewVersion+" can't find.")
	os.chdir(CWD)

print("==============Fv File Rename And Copy To Pkg============================================")# This block can modify.
for Match in range(len(Matchlist)):# Fv File Rename And Copy To Pkg
	if Matchlist[Match].split("_")[1] in Processlist:
		if (os.path.isdir(CWD+"\\"+Matchlist[Match])):# Check Fv Folder Exist
			Path=CWD+"\\"+Matchlist[Match]
			os.chdir(Path)
			if "Fv_Q23_"+NewVersion+"_" in Matchlist[Match]:# For 16MB BIOS
			# If add 16MB BIOS Please add [or "Fv_QXX_"+NewVersion+"_" in Matchlist[Match]] in it
				for Boardnum in range(len(BoardID)):# Find which BoardID is right
					Boardversion=BoardID[Boardnum]+"_"+NewVersion
					if os.path.isfile(Path+"\\"+Boardversion+"_12.bin") and os.path.isfile(Path+"\\"+Boardversion+".xml"):# If Alreadly renamed
						print(Boardversion+"_12.bin & "+Boardversion+".xml alreadly renamed.")
					if os.path.isfile(Path+"\\"+Boardversion+".bin") and os.path.isfile(Path+"\\"+BoardID[Boardnum]+".xml"):
						os.rename(Path+"\\"+Boardversion+".bin", Path+"\\"+Boardversion+"_12.bin")# Rename Fv folder 2 files
						os.rename(Path+"\\"+BoardID[Boardnum]+".xml", Path+"\\"+Boardversion+".xml")
						print(Boardversion+"_12.bin & "+Boardversion+".xml rename succeeded.")
					if os.path.isfile(Path+"\\"+Boardversion+".inf"):
						if NewBuildID=="" or NewBuildID=="0000":
							if os.path.isdir(CWD+"\\"+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion):# Check Pkg Folder Exist
								CopyFiles(Matchlist[Match], Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion)
							else:	print("Pkg "+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+" can't find.")
						else:
							if os.path.isdir(CWD+"\\"+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+"_"+NewBuildID):# Check Pkg Folder Exist
								CopyFiles(Matchlist[Match], Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+"_"+NewBuildID)
							else:	print("Pkg "+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+"_"+NewBuildID+" can't find.")
				os.chdir(CWD)
			else:
				for Boardnum in range(len(BoardID)):# Find which BoardID is right
					Boardversion=BoardID[Boardnum]+"_"+NewVersion
					if os.path.isfile(Path+"\\"+Boardversion+"_16.bin") and os.path.isfile(Path+"\\"+Boardversion+".xml"):# If Alreadly renamed
						print(Boardversion+"_16.bin & "+Boardversion+".xml alreadly renamed.")
					if os.path.isfile(Path+"\\"+Boardversion+".bin") and os.path.isfile(Path+"\\"+BoardID[Boardnum]+".xml"):
						os.rename(Path+"\\"+Boardversion+".bin", Path+"\\"+Boardversion+"_16.bin")# Rename Fv folder 2 files
						os.rename(Path+"\\"+BoardID[Boardnum]+".xml", Path+"\\"+Boardversion+".xml")
						print(Boardversion+"_16.bin & "+Boardversion+".xml rename succeeded.")
					if os.path.isfile(Path+"\\"+Boardversion+".inf"):
						if NewBuildID=="" or NewBuildID=="0000":
							if os.path.isdir(CWD+"\\"+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion):# Check Pkg Folder Exist
								CopyFiles(Matchlist[Match], Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion)
							else:	print("Pkg "+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+" can't find.")
						else:
							if os.path.isdir(CWD+"\\"+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+"_"+NewBuildID):# Check Pkg Folder Exist
								CopyFiles(Matchlist[Match], Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+"_"+NewBuildID)
							else:	print("Pkg "+Projectlist[Boardnum]+"_"+Phase+"_"+Boardversion+"_"+NewBuildID+" can't find.")
				os.chdir(CWD)
	else:	print("Need to be processed Fv folder:"+Matchlist[Match]+" can't find.")
if len(Matchlist)==0 or len(Processlist)==0:	print("Can't find anything Fv folder.")

print("========================================================================================")
CheckPkg()# Check new release Pkg is OK?
print("\nFinally pkg please compare with leading project.")
sys.exit()
# Script End 
