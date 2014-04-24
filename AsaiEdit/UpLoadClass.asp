<%
'****************  风声无组件上传类 2.11 冷风加强版说明 *****************
'在风声无组件上传类2.11基础上增加了两个实用的功能：
'功能一：多级目录创建函数:CreateFolders(Folders)——根据SavePath定义的多级目录自动创建
'功能二：恶意文件检查函数：CheckFile(File_path)——检查恶意文件，发现恶意文件则删除上传的文件，恶意特证码可在函数中自定义
'时间：2008-4-5
'**********************************************************
'----------------------------------------------------------
Class UpLoadClass
	Private m_TotalSize,m_MaxSize,m_FileType,m_SavePath,m_AutoSave,m_Error,m_Charset
	Private m_dicForm,m_binForm,m_binItem,m_strDate,m_lngTime
	Public	FormItem,FileItem

	Public Property Get Version
		Version="Fonshen UpLoadClass Version 2.11"
	End Property

	Public Property Get Error
		Error=m_Error
	End Property

	Public Property Get Charset
		Charset=m_Charset
	End Property
	Public Property Let Charset(strCharset)
		m_Charset=strCharset
	End Property

	Public Property Get TotalSize
		TotalSize=m_TotalSize
	End Property
	Public Property Let TotalSize(lngSize)
		if isNumeric(lngSize) then m_TotalSize=Clng(lngSize)
	End Property

	Public Property Get MaxSize
		MaxSize=m_MaxSize
	End Property
	Public Property Let MaxSize(lngSize)
		if isNumeric(lngSize) then m_MaxSize=Clng(lngSize)
	End Property

	Public Property Get FileType
		FileType=m_FileType
	End Property
	Public Property Let FileType(strType)
		m_FileType=strType
	End Property

	Public Property Get SavePath
		SavePath=m_SavePath
	End Property
	Public Property Let SavePath(strPath)
		m_SavePath=Replace(strPath,chr(0),"")
	End Property

	Public Property Get AutoSave
		AutoSave=m_AutoSave
	End Property
	Public Property Let AutoSave(byVal Flag)
		select case Flag
			case 0,1,2: m_AutoSave=Flag
		end select
	End Property

	Private Sub Class_Initialize
		m_Error	   = -1
		m_Charset  = "gb2312"
		m_TotalSize= 0
		m_MaxSize  = 3072000
		m_FileType = "png/jpg/gif/doc/docx/xls/xlsx/dwg"
		m_SavePath = ""
		m_AutoSave = 0
		Dim dtmNow : dtmNow = Date()
'		m_strDate  = Year(dtmNow)&Right("0"&Month(dtmNow),2)&Right("0"&Day(dtmNow),2)
'		m_lngTime  = Clng(Timer()*1000)
		Set m_binForm = Server.CreateObject("ADODB.Stream")
		Set m_binItem = Server.CreateObject("ADODB.Stream")
		Set m_dicForm = Server.CreateObject("Scripting.Dictionary")
		m_dicForm.CompareMode = 1
	End Sub

	Private Sub Class_Terminate
		m_dicForm.RemoveAll
		Set m_dicForm = nothing
		Set m_binItem = nothing
		m_binForm.Close()
		Set m_binForm = nothing
	End Sub

	Public Function Open()
		
		Open = 0
		if m_Error=-1 then
			m_Error=0
		else
			Exit Function
		end if
		Dim lngRequestSize : lngRequestSize=Request.TotalBytes
		if m_TotalSize>0 and lngRequestSize>m_TotalSize then
			m_Error=5
			Exit Function
		elseif lngRequestSize<1 then
			m_Error=4
			Exit Function
		end if

		Dim lngChunkByte : lngChunkByte = 102400
		Dim lngReadSize : lngReadSize = 0
		m_binForm.Type = 1
		m_binForm.Open()
		do
			m_binForm.Write Request.BinaryRead(lngChunkByte)
			lngReadSize=lngReadSize+lngChunkByte
			if  lngReadSize >= lngRequestSize then exit do
		loop		
		m_binForm.Position=0
		Dim binRequestData : binRequestData=m_binForm.Read()

		Dim bCrLf,strSeparator,intSeparator
		bCrLf=ChrB(13)&ChrB(10)
		intSeparator=InstrB(1,binRequestData,bCrLf)-1
		strSeparator=LeftB(binRequestData,intSeparator)

		Dim strItem,strInam,strFtyp,strPuri,strFnam,strFext,lngFsiz
		Const strSplit="'"">"
		Dim strFormItem,strFileItem,intTemp,strTemp
		Dim p_start : p_start=intSeparator+2
		Dim p_end
		Do
			p_end = InStrB(p_start,binRequestData,bCrLf&bCrLf)-1
			m_binItem.Type=1
			m_binItem.Open()
			m_binForm.Position=p_start
			m_binForm.CopyTo m_binItem,p_end-p_start
			m_binItem.Position=0
			m_binItem.Type=2
			m_binItem.Charset=m_Charset
			strItem = m_binItem.ReadText()
			m_binItem.Close()
			intTemp=Instr(39,strItem,"""")
			strInam=Mid(strItem,39,intTemp-39)

			p_start = p_end + 4
			p_end = InStrB(p_start,binRequestData,strSeparator)-1
			m_binItem.Type=1
			m_binItem.Open()
			m_binForm.Position=p_start
			lngFsiz=p_end-p_start-2
			m_binForm.CopyTo m_binItem,lngFsiz

			if Instr(intTemp,strItem,"filename=""")<>0 then
			if not m_dicForm.Exists(strInam&"_From") then
				strFileItem=strFileItem&strSplit&strInam
				if m_binItem.Size<>0 then
					intTemp=intTemp+13
					strFtyp=Mid(strItem,Instr(intTemp,strItem,"Content-Type: ")+14)
					strPuri=Mid(strItem,intTemp,Instr(intTemp,strItem,"""")-intTemp)
					intTemp=InstrRev(strPuri,"\")
					strFnam=Mid(strPuri,intTemp+1)
					m_dicForm.Add strInam&"_Type",strFtyp
					m_dicForm.Add strInam&"_Name",strFnam
					m_dicForm.Add strInam&"_Path",Left(strPuri,intTemp)
					m_dicForm.Add strInam&"_Size",lngFsiz
					if Instr(strFnam,".")<>0 then
						strFext=Mid(strFnam,InstrRev(strFnam,".")+1)
					else
						strFext=""
					end if

					select case strFtyp
					case "image/jpeg","image/pjpeg","image/jpg"
						if Lcase(strFext)<>"jpg" then strFext="jpg"
						m_binItem.Position=3
						do while not m_binItem.EOS
							do
								intTemp = Ascb(m_binItem.Read(1))
							loop while intTemp = 255 and not m_binItem.EOS
							if intTemp < 192 or intTemp > 195 then
								m_binItem.read(Bin2Val(m_binItem.Read(2))-2)
							else
								Exit do
							end if
							do
								intTemp = Ascb(m_binItem.Read(1))
							loop while intTemp < 255 and not m_binItem.EOS
						loop
						m_binItem.Read(3)
						m_dicForm.Add strInam&"_Height",Bin2Val(m_binItem.Read(2))
						m_dicForm.Add strInam&"_Width",Bin2Val(m_binItem.Read(2))
					case "application/x-shockwave-flash"
						if Lcase(strFext)<>"swf" then strFext="swf"
						m_binItem.Position=0
						if Ascb(m_binItem.Read(1))=70 then
							m_binItem.Position=8
							strTemp = Num2Str(Ascb(m_binItem.Read(1)), 2 ,8)
							intTemp = Str2Num(Left(strTemp, 5), 2)
							strTemp = Mid(strTemp, 6)
							while (Len(strTemp) < intTemp * 4)
								strTemp = strTemp & Num2Str(Ascb(m_binItem.Read(1)), 2 ,8)
							wend
							m_dicForm.Add strInam&"_Width", Int(Abs(Str2Num(Mid(strTemp, intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 1, intTemp), 2)) / 20)
							m_dicForm.Add strInam&"_Height",Int(Abs(Str2Num(Mid(strTemp, 3 * intTemp + 1, intTemp), 2) - Str2Num(Mid(strTemp, 2 * intTemp + 1, intTemp), 2)) / 20)
						end if
					end select

					m_dicForm.Add strInam&"_Ext",strFext
					m_dicForm.Add strInam&"_From",p_start
					if m_AutoSave<>2 then
						intTemp=GetFerr(lngFsiz,strFext)
						m_dicForm.Add strInam&"_Err",intTemp
						if intTemp=0 then
							if m_AutoSave=0 then
								strFnam=GetTimeStr()
								if strFext<>"" then strFnam=strFnam&"."&strFext
							end if
							CreateFolders(m_SavePath)'创建目录
							dim upload_path:upload_path=Server.MapPath(m_SavePath&strFnam)	'文件保存路径
							m_binItem.SaveToFile upload_path,2	'保存文件
							m_dicForm.Add strInam,strFnam
						end if
					end if
				else
					m_dicForm.Add strInam&"_Err",-1
				end if
			end if
			else
				m_binItem.Position=0
				m_binItem.Type=2
				m_binItem.Charset=m_Charset
				strTemp=m_binItem.ReadText
				if m_dicForm.Exists(strInam) then
					m_dicForm(strInam) = m_dicForm(strInam)&","&strTemp
				else
					strFormItem=strFormItem&strSplit&strInam
					m_dicForm.Add strInam,strTemp
				end if
			end if

			m_binItem.Close()
			p_start = p_end+intSeparator+2
		loop Until p_start+3>lngRequestSize
		FormItem=Split(strFormItem,strSplit)
		FileItem=Split(strFileItem,strSplit)
		if CheckFile(upload_path) then m_dicForm.Add strInam&"_Err",4	'恶意文件，返回错误码：4
		Open = lngRequestSize
	End Function

	Private Function GetTimeStr()
'		m_lngTime=m_lngTime+1
'		GetTimeStr=m_strDate&Right("00000000"&m_lngTime,8)
dim ranNum 
dim dtNow 
dtNow=Now() 
randomize 
ranNum=int(90000*rnd)+10000 
'以下这段由webboy提供 
GetTimeStr=year(dtNow) & right("0" & month(dtNow),2) & right("0" & day(dtNow),2) & right("0" & hour(dtNow),2) & right("0" & minute(dtNow),2) & right("0" & second(dtNow),2) & ranNum 	
	End Function

	Private Function GetFerr(lngFsiz,strFext)
		dim intFerr
		intFerr=0
		if lngFsiz>m_MaxSize and m_MaxSize>0 then
			if m_Error=0 or m_Error=2 then m_Error=m_Error+1
			intFerr=intFerr+1
		end if
		if Instr(1,LCase("/"&m_FileType&"/"),LCase("/"&strFext&"/"))=0 and m_FileType<>"" then
			if m_Error<2 then m_Error=m_Error+2
			intFerr=intFerr+2
		end if
		GetFerr=intFerr
	End Function

	Public Function Save(Item,strFnam)'当设置自定义保存文件名时，调用此函数保存文件
		Save=false
		if m_dicForm.Exists(Item&"_From") then
			dim intFerr,strFext
			strFext=m_dicForm(Item&"_Ext")
			intFerr=GetFerr(m_dicForm(Item&"_Size"),strFext)
			if m_dicForm.Exists(Item&"_Err") then
				if intFerr=0 then
					m_dicForm(Item&"_Err")=0
				end if
			else
				m_dicForm.Add Item&"_Err",intFerr
			end if
			if intFerr<>0 then Exit Function
			if VarType(strFnam)=2 then
				select case strFnam
					case 0:strFnam=GetTimeStr()
						if strFext<>"" then strFnam=strFnam&"."&strFext
					case 1:strFnam=m_dicForm(Item&"_Name")
				end select
			end if
			m_binItem.Type = 1
			m_binItem.Open
			m_binForm.Position = m_dicForm(Item&"_From")
			m_binForm.CopyTo m_binItem,m_dicForm(Item&"_Size")
			CreateFolders(m_SavePath)'创建目录
			dim upload_path:upload_path=Server.MapPath(m_SavePath&strFnam)	'文件保存路径
			m_binItem.SaveToFile upload_path,2
			m_binItem.Close()
			if m_dicForm.Exists(Item) then
				m_dicForm(Item)=strFnam
			else
				m_dicForm.Add Item,strFnam
			end if
			if CheckFile(upload_path) then
				 m_dicForm(Item&"_Err")=4	'恶意文件，返回错误码：4
			else
				Save=true
			end if
		end if
	End Function

	Public Function GetData(Item)
		GetData=""
		if m_dicForm.Exists(Item&"_From") then
			if GetFerr(m_dicForm(Item&"_Size"),m_dicForm(Item&"_Ext"))<>0 then Exit Function
			m_binForm.Position = m_dicForm(Item&"_From")
			response.write m_binForm.Read(m_dicForm(Item&"_Size"))
			GetData = m_binForm.Read(m_dicForm(Item&"_Size"))
		end if
	End Function

	Public Function Form(Item)
		if m_dicForm.Exists(Item) then
			Form=m_dicForm(Item)
		else
			Form=""
		end if
	End Function

	Private Function BinVal2(bin)
		dim lngValue,i
		lngValue = 0
		for i = lenb(bin) to 1 step -1
			lngValue = lngValue *256 + Ascb(midb(bin,i,1))
		next
		BinVal2=lngValue
	End Function

	Private Function Bin2Val(bin)
		dim lngValue,i
		lngValue = 0
		for i = 1 to lenb(bin)
			lngValue = lngValue *256 + Ascb(midb(bin,i,1))
		next
		Bin2Val=lngValue
	End Function

	Private Function Num2Str(num, base, lens)
		Dim ret,i
		ret = ""
		while(num >= base)
			i   = num Mod base
			ret = i & ret
			num = (num - i) / base
		wend
		Num2Str = Right(String(lens, "0") & num & ret, lens)
	End Function

	Private Function Str2Num(str, base)
		Dim ret, i
		ret = 0 
		for i = 1 to Len(str)
			ret = ret * base + Cint(Mid(str, i, 1))
		next
		Str2Num = ret
	End Function
	
	'========================
	'功能：创建多级目录函数
	'时间：2008-4-5
	private Function CreateFolders(Folders)
		dim F_str,n,temp,F_path
		set FSO=server.CreateObject("Scripting.FileSystemObject")
		F_str=replace(replace(trim(Folders)," ","_"),"\","/")	'将目录中的空格转为下划线
		F_arr=split(F_str,"/")
		temp=""
		for n=0 to ubound(F_arr)
			temp=temp&F_arr(n)&"/"
			if F_arr(n)<>".." and F_arr(n)<>"" then
				F_path=server.mapPath(temp)
				if not FSO.FolderExists(F_path) then
					FSO.CreateFolder F_path
				end if
			end if
		next
		set FSO=nothing
	end Function
	
	'========================
	'功能：恶意文件检查
	'返回值：true==文件含有恶意代码；false==文件正常
	'时间：2008-4-5
	Function CheckFile(File_path)
		dim objFSO,objTS,strtext,ComStr,strArray
		checkFile=false
		Set objFSO=Server.CreateObject("Scripting.FileSystemObject")
		If objFSO.FileExists(File_path) Then
			Set objTS=objFSO.OpenTextFile(File_path,1)
			strText=lcase(objTS.ReadAll)
			objTS.Close
				'禁止字符，可随时添加
			ComStr="sj-fans|cookie|.getfolder|.createfolder|.deletefolder|.createdirectory|.deletedirectory|0n error resume next|adodb.stream|createobject|scripting.filesystemobject|strbackdoor|password|command.com"
			ComStr=ComStr&"|.saveas|wscript.shell|shell.application|script.encode|folderpath|session|request|iframe|frame|execute|object|server.mappath|insert into|Conn,1,|.update|Response" 
			strArray=split(ComStr,"|")
			for i=0 to ubound(strArray)
				if instr(strText,strArray(i))<>0 then
					objFSO.DeleteFile File_path,True	'如果文件含有恶意代码，则删除文件
					CheckFile=true
					exit for
				end if
			next
		end if
		Set objFSO=nothing
	end Function
End Class
%>