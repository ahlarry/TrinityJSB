<%
Class CLSXJWEB
	Private strCode, i, IsConn
	Public OpDBNum
	Private Sub Class_Initialize()
		strCode="" : i=0 : IsConn=False : OpDBNum=0
		'���
	End Sub

	Private Sub Class_Terminate()
		'���
	End Sub

	Public Sub Header()
%><!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="NO-CACHE">
<meta http-equiv="Expires" content="Wed, 26 Feb 1997 08:21:57 GMT">
<meta name="Generator" content="EditPlus">
<meta name="Author" content="<%=web_info(3)%>">
<meta name="Description" content="<%=web_info(4)%>">
<meta name="Keywords" content="<%=web_info(5)%>">
<%
		If CssFiles<>"" Then
			CssFiles=Split(CssFiles, "||")
			for i=0 to ubound(CssFiles)
%><link href="<%=web_info(2) & CssFiles(i)%>" rel="stylesheet" type="text/css">
<%
			next
		End If

		If JsFiles<>"" Then
			JsFiles=Split(JsFiles, "||")
			for i=0 to ubound(JsFiles)
%><script language="javascript" src="<%=web_info(2) & JsFiles(i)%>"></script>
<%
			next
		End If
%><title><%=web_info(0)%> �� <%=CurPage%></title>
<%
	Response.write(Comm_Css)
%>
</head>
<%
		Response.Write("<body")
		If web_info(6) Then Response.Write(" oncopy='return false;' onselectstart='return false;' oncontextmenu='return false;'")
	    If web_info(7)<>"" Then Response.Write(" onmouseover=""window.status='" & web_info(7) & "'; return true;""")
		Response.Write(">")
%>
<%
	End Sub

	Public Sub Footer()
%>
</body>
</html>
<%
	End Sub

	'-------------------------------�򿪡��������ݿ�--------------------------------------
	Private Sub DBConnect()
		'On Error Resume Next
		Set Conn=Server.CreateObject("ADODB.Connection")
		Set Rs=Server.CreateObject("ADODB.Recordset")
		On error resume next
		Conn.Open ConnStr
		If Err Then
			err.Clear
			Set Rs=Nothing
			Set Conn = Nothing
			Response.Write "<br><br><font style=font-size:16pt;>���ݿ�����ά����! ���Ժ���ʣ�лл��</font>"
			Response.End
		End If
		IsConn=True
	End Sub

	Public Function Exec(sql,etype)
		If IsConn=False Then Call DBConnect()
		Select Case etype
			Case 0
				Conn.Execute(sql)
			Case 1
				Set Exec=Conn.Execute(sql)
			Case else
				'Conn.Execute(sql)
		End Select
		OpDBNum=OpDBNum+1
	End Function

	'��¼����
	Function RsCount(strtb)		'�鿴��¼��
		Dim tmpRs
		Set tmpRs=Server.CreateObject("ADODB.Recordset")
		Set tmpRs=Exec("select count(*) from " & strtb, 1)
		If IsNull(tmpRs(0)) then
			RsCount=0
		Else
			RsCount = tmpRs(0)
		End If
		tmpRs.Close
		Set tmpRs=Nothing
	End Function

	'-------------------------------�жϷ����Ƿ������ⲿ-------------------------------
	Private Function ChkPost()
		Dim Server_v1, Server_v2
		Server_v1=Request.ServerVariables("HTTP_REFERER")
		Server_v2=Request.ServerVariables("SERVER_NAME")
		If Server_v1<>"" Then
			Server_v2=LCase("http://"&server_v2)
			If LCase(Left(Server_v1,Len(Server_v2)))=Server_v2 then
				ChkPost=True
				Exit Function
			End If
		End If
		ChkPost=False
	End Function

	Public Function SubmitChk()
		SubmitChk=ChkPost()
	End Function

	Public Function Chk()
		Chk=False
		If Trim(Request("chk"))="yes" Then
			Chk=ChkPost()
		End If
	End Function

	'-------------------------------����ת��-------------------------------------------------
	Public Function HtmlToCode(str)			'Ϊ�˽�HTML������ҳ������ʾ
		If str="" Or IsNull(str) Then HtmlToCode="&nbsp; " : Exit Function
		str = Replace(str, ">", "&gt;")
		str = Replace(str, "<", "&lt;")
		str = Replace(str, Chr(32), "&nbsp;")
		str = Replace(str, Chr(9), "&nbsp;")
		str = Replace(str, Chr(34), "&quot;")
		str = Replace(str, Chr(39), "&#39;")		'������
		str = Replace(str, Chr(13), "")
		str = Replace(str, Chr(10) & Chr(10), "</p><p>")
		str = Replace(str, Chr(10), "<br>")
		HtmlToCode = str
	End Function

	Public Function CodeToHtml(str)		'��������ΪHTML����
		If IsNull(str) Then CodeToHtml="" : Exit Function
		'str = Replace(str, "&nbsp", " ")
		'str = Replace(str, "<br>", vbcrlf)
		str = Replace(str, " ", "&nbsp")
		str = Replace(str, vbcrlf, "<br>")
		CodeToHtml = str
	End Function

	Public Function UserIP(iagent)		'iagent=0��ʾIP��ַ������Ϊ�Ƿ����
		Dim ip1,ip2
		ip1=request.servervariables("http_x_forwarded_for")
		ip2=request.servervariables("remote_addr")
		if instr(ip1,",")>0 then ip1=left(ip1,instr(ip1,",")-1)
		if instr(ip2,",")>0 then ip2=left(ip2,instr(ip2,",")-1)
		if ip1 <> "" then
			if iagent = 0 then
				userip = ip1
			else
				userip = true
			end if
		else
			if iagent = 0 then
				userip = ip2
			else
				userip = false
			end if
		end if
	end function

	Public Function UserSys(i)	'i=0 Ϊ����ϵͳ 1Ϊ�����
		dim browser, version, platform, agent
		browser = "unknown"
		version = "unknown"
		platform = "unknown"
		agent = lcase(request.servervariables("http_user_agent"))
		agent = split(agent,";")
		if instr(agent(1),"msie") > 0 then
			browser = "MS IE"
			version = trim(left(replace(agent(1),"msie",""),6))
			browser = browser & version
			'if instr(agent(3),"tencenttraveler") > 0 then browser = browser & " (TencentTraveler)"
			'if instr(agent(3),"myie") > 0 then browser = browser & "(MYIE)"
		elseif instr(agent(4), "netscape") > 0 then
			browser = "netscape"
			version = split(agent(4),"/")
			browser = browser & version
		end if
		if instr(agent(2), "nt 5.2") > 0 then
			platform = "Windows 2003 Server"
		elseif instr(agent(2), "nt 5.1") > 0 then
			platform = "Windows XP"
		elseif instr(agent(2), "nt 5.0") > 0 then
			platform = "Windows 2000"
		elseif instr(agent(2), "9x") > 0 then
			platform = "Windows ME"
		elseif instr(agent(2), "98") > 0 then
			platform = "Windows 98"
		elseif instr(agent(2), "95") > 0 then
			platform = "Windows 95"
		end if
		select case i
			case 0
				usersys = platform
			case 1
				usersys = browser
			case else
				usersys = "�����:" & browser &  " ����ϵͳ:" & platform
		end select
	End function

	Rem �ַ��г�
	Function StringCut(fstring,num)
		Dim ctypes,cnum,ci,tt,tc,cc,cmod
		cmod=3
		ctypes=fstring
		cnum=cint(num)
		StringCut=""
		tc=0
		cc=0
		for ci=1 to len(ctypes)
			if cnum<0 then
				StringCut=StringCut&"..."
				exit for
			end if
			tt=mid(ctypes,ci,1)
			if int(asc(tt))>=0 then
				StringCut=StringCut&tt
				tc=tc+1
				cc=cc+1
				if tc=2 then tc=0
				cnum=cnum-1
				if cc>cmod then cnum=cnum-1
				cc=0
			else
				cnum=cnum-1
				if cnum<=0 then
					StringCut=StringCut&"..."
					exit for
				end if
				StringCut=StringCut&tt
			end if
		next
	End Function

  '-----------------------------------���Nλ����-------------------------------------
	Function Rand_Num(rnum)
		dim ri,rmax,rmin,rndnum
		rmax=10^(rnum)-1
		rmin=10^(rnum-1)
		randomize
		rndnum=int((rmax-rmin+1)*rnd)+rmin
		for ri=1 to rnum-len(rndnum)
			rndnum="0"&rndnum
		next
		rand_num=rndnum
	End Function
End Class


Class CLSDBOP
	Private strSql, objAdoxDatabase
	Private Sub Class_Initialize()
		strSql=""
		'���
	End Sub

	Private Sub Class_Terminate()
		strSql=""
		Set objAdoxDatabase=Nothing
	End Sub

	Public Sub CreateTable(tablename)	'�½���
		On Error Resume Next
		strSql="Create table "&tablename&" (id int identity (1, 1) not null)"
		Call Conn.Execute(strSql)
		If Err Then
			Response.Write "<br>�½��� "& tablename &" <font color=blue>ʧ��</font>��ԭ��" & Err.Description & "<br>"
			Err.Clear
			Response.Flush
		Else
			Response.Write "<br>�½��� <b>"&tablename&"</b> �ɹ�<br>"
			Response.Flush
		End If
	End Sub

	'�������ݿ��������ڲ������ϱ������±���
	Public Sub RenameTable(oldname, newname)
		On Error resume Next
		Set objAdoxDatabase = Server.CreateObject("ADOX.Catalog")
		objAdoxDatabase.ActiveConnection = Connstr
		If Err Then
			Response.Write "<br>�������ı��������������Ҫ�����Ŀռ䲻֧�ִ˶������ܿ�����Ҫ�ֶ����ı�����ԭ��" & Err.Description & "<br>"
			Err.Clear
			Response.Flush
		End If

		objAdoxDatabase.tables(oldname).name = newname
		If Err Then
			Response.Write "<br>���ı���<font color=blue>����</font>�����ֶ������ݿ��� <b>"&oldname&"</b> ��������Ϊ <b>"&newname&"</b>��ԭ��" & Err.Description & "<br>"
			Err.Clear
			Response.Flush
		Else
			Response.Write "<br>���ı��� "&oldname&" to "&newname&" �ɹ� <br>"
			Response.Flush
		End IF
	End Sub

	'ɾ���ֶ�ͨ�ú���
	Public Sub DelColumn(tablename,columnname)
		On Error Resume Next
		strSql = "alter table "&tablename&" drop "&columnname&""
		Call Conn.Execute(strsql)
		If Err Then
			Response.Write "<br>ɾ�� "&tablename&" �����ֶ�<font color=blue>����</font>�����ֶ������ݿ��� <b>"&columnname&"</b> �ֶ�ɾ����ԭ��" & Err.Description & "<br>"
			Err.Clear
			Response.Flush
		Else
			Response.Write "<br>ɾ�� "&tablename&" �����ֶ� "&columnname&" �ɹ� <br>"
			Response.Flush
		End IF
	End Sub

	'����ֶ�ͨ�ú���
	Public Sub AddColumn(tablename,columnname,columntype)
		On Error Resume Next
		strSql = "alter table "&tablename&" add "&columnname&" "&columntype&""
		Call Conn.Execute(strSql)
		If Err Then
			Response.Write "<br>�½� "&tablename&" �����ֶ�<font color=blue>����</font>�����ֶ������ݿ��� <b>"&columnname&"</b> �ֶν���������Ϊ <b>"&columntype&"</b>��ԭ��" & err.description & "<br>"
			Err.Clear
			Response.Flush
		Else
			Response.Write "<br>�½� "&tablename&" �����ֶ� "&columnname&" �ɹ� <br>"
			Response.Flush
		End If
	End Sub

	'�����ֶ�����ͨ�ú���
	Public Sub ModColumn(tablename,columnname,columntype)
		On Error Resume Next
		strSql = "alter table "&tablename&" alter column "&columnname&" "&columntype&""
		Call Conn.Execute(strsql)
		If Err Then
			Response.Write "<br>���� "&tablename&" �����ֶ�����<font color=blue>����</font>�����ֶ������ݿ��� <b>"&columnname&"</b> �ֶθ���Ϊ <b>"&columntype&"</b> ���ԣ�ԭ��" & err.description & "<br>"
			Err.Clear
			Response.Flush
		Else
			Response.Write "<br>���� "&tablename&" �����ֶ����� "&columnname&" �ɹ� <br>"
			Response.Flush
		End If
	End Sub

	'ɾ����
	Public Sub DelTable(tablename)
		On Error Resume Next
		strSql = "drop table "&tablename&""
		Call Conn.Execute(strsql)
		If Err Then
			Response.Write "<br>ɾ�� <b>"&tablename&"</b> ��<font color=blue>����</font>��ԭ��" & Err.Description & "  �����"&tablename&"����,���ֶ������ݿ��� <b>"&tablename&"</b> ��ɾ��"
			Err.Clear
			Response.Flush
		Else
			Response.Write "<br>ɾ�� <b>"&tablename&"</b> ��ɹ� <br>"
			Response.Flush
		End If
	End Sub
End Class
%>