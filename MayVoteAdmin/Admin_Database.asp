<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%'����Ա��֤
Call CheckUnAdmin()
'��ֹ�ǳ�������Ա����
Call CheckUnAdmin1()
%>
<html>
<head>
<title>���ݿ����</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#F6F6F6" leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>
<br>
<table width='99%' border='1' align='center' cellpadding='3' cellspacing='0' bordercolor="#666666" >
  <tr> 
    <td height='25'  align='center'  background="Images/title.gif"><strong><font color="#FFFFFF">�� 
      �� �� �� ��</font></strong></td>
  </tr>
  <tr > 
    <td height='30' bgcolor="#FFF3E6" >&nbsp;&nbsp;<strong>��������</strong> <a href='Admin_Database.asp?Action=BackUpData' target=main>�������ݿ�</a>&nbsp;|&nbsp;<a href='Admin_Database.asp?Action=RestoreData' target=main>�ָ����ݿ�</a>&nbsp;|&nbsp;<a href='Admin_Database.asp?Action=CompactData' target=main>ѹ�����ݿ�</a></td>
  </tr>
</table>
<br>
<%
		Action=trim(Request("Action"))
		If Action="BackUpData" then
				Call BackUpData()
		ElseIf Action="RestoreData" then
				Call RestoreData()
		ElseIf Action="CompactData" then
				Call CompactData()
		ElseIf Action="BackUpDataYes" then
				Call BackUpDataYes()
		ElseIf Action="RestoreDataYes" then
				Call RestoreDataYes()
		ElseIf Action="CompactDataYes" then
				Call CompactDataYes()
		Else
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ�������Action=<font color=red>"&Action&"</font></li>&Action=OtherErr"
		End If
'************
'*  ����ѡ��  *
'************
		Sub BackUpData()
%>
<form action='Admin_Database.asp?Action=BackUpDataYes' method='post' name="backup" id="backup" target="main">
  <table width='99%' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#666666" >
    <tr>
      <td height='25' align='center' valign='middle' background="Images/title.gif"> 
        <b><font color="#FFFFFF">�� �� �� �� ��</font></b> </td> 
</tr>
<tr >
      <td height='150' align='center' valign='middle' bgcolor="#FFF3E6"> 
        <table width='99%' border='1' cellpadding='3' cellspacing='0' bordercolor="#FFFFFF">
          <tr> 
            <td width=30% height='34' align='right'>��ǰ���ݿ�·����</td>
            <td width=30%> 
              <input type="hidden" size="20" name="DBpath" value="<%=Replace(AdminDbPath,"/","\") &Replace(mdb,"/","\") %>"> 
              <%=AdminDbPath&mdb%></td>
            <td width=40%>���ݿ�����Ŀ¼�����·��</td>
          </tr>
          <tr> 
            <td height='34' align='right'>����Ŀ¼��</td>
            <td><input type="text" size="20" name="bkfolder" value="Databackup" readonly> </td>
            <td>���·��Ŀ¼����Ŀ¼�����ڣ����Զ�����<br>Ĭ��·������ò�Ҫ����</td>
          </tr>
          <tr> 
            <td height='34' align='right'>�������ƣ�</td>
            <td height='34'><input type="text" size="20" name="bkDBname" value='MayVote_backup'> 
            </td>
            <td height='34'>����ͬ���ļ���������<br>
              <font color="#FF0000">�ļ���׺ǿ��Ϊ .asa ,�벻Ҫ����д�ļ���׺</font></td>
          </tr>
          <tr> 
            <td height='20' colspan='3' align="center"> 
              <input name='Submit' type="Submit" value=' ��ʼ���� '> 
              <input name="Action" type="hidden" id="Action" value="BackUpYes"></td>
          </tr>
          <tr> 
            <td height='20' colspan='3'>( �������ݿ���ҪFSO֧�֣�FSO��ذ����뿴΢����վ ) <br>
              -----------------------------------------------------------------------------------------<br>
              ��������д����������ݿ�·��ȫ�����������Ĭ�����ݿ��ļ�Ϊ../Data/MayVote.mdb����һ��������Ĭ�����������������ݿ�<br>
              ���������������������������̳���ݣ��Ա�֤�������ݰ�ȫ��<br>
              ע�⣺����·��������������ռ��Ŀ¼�����·�� </td>
          </tr>
        </table>
</td>
</tr>
</table>
</form>
<%End Sub

'************
'*��ԭѡ��    *
'************
Sub RestoreData()
%>
<form action="Admin_Database.asp?Action=RestoreDataYes" method='post' name="RestoreYes" id="RestoreYes" target="main">
  <table width='99%' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#666666" >
    <tr class='title'> 
      <td height='25' align='center' valign='middle' background="Images/title.gif"><b><font color="#FFFFFF">�� 
        �� �� �� ��</font></b></td>
  </tr>
    <tr > 
      <td bgcolor="#FFF3E6"> 
        <table width='99%' border='1' cellpadding='3' cellspacing='0' bordercolor="#FFFFFF">
          <tr> 
            <td width='30%' height='30' align='right'>�������ݿ�·������ԣ���</td>
            <td width="70%" height='30' align="left"> 
              <input name="dbpath" type=text id="dbpath" value="Databackup\MayVote_Backup" size="50" maxlength="200">
              <strong><font color="#FF0000"> </font></strong></td>
          </tr>
          <tr> 
            <td height='20' colspan='2' align="center"> 
              <input name='Submit' type="Submit" value=" �ָ����� "> 
              <input name="Action" type="hidden" id="Action" value="RestoreYes"> 
            </td>
          </tr>
          <tr> 
            <td height='20' colspan='2'>��д����ǰʹ�õ����ݿ�·�����粻�븲�ǵ�ǰ�ļ���������������ע��·���Ƿ���ȷ����Ȼ���޸�Conn.asp�ļ������Ŀ���ļ����͵�ǰʹ�����ݿ���һ�µĻ��������޸�Conn.asp�ļ�<br>
              -----------------------------------------------------------------------------------------<br>
              ��������д����������ݿ�·��ȫ�����������Ĭ�ϱ������ݿ��ļ�ΪDataBackup/MayVote_Backup.asa���밴�����ı����ļ������޸ġ�<br>
              ������������������ָ�������̳���ݣ��Ա�֤�������ݰ�ȫ��<br>
              ע�⣺1.����·��������������ռ��Ŀ¼�����·�� <br>
              ������2.<font color="#FF0000">���ݿⱸ���ļ���׺�Ѿ�ǿ��Ϊ.asa</font> </td>
          </tr>
        </table>
        
      </td>
  </tr>
</table>
</form>
<%
End Sub

'**************
'*���ݿ��ѹ��*
'**************
Sub CompactData()
%>
<form action="Admin_Database.asp?Action=CompactDataYes" method='post' name="Compact" id="Compact" target="main">
  <table width='99%' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#666666" >
    <tr class='title'> 
      <td height='25' align='center' valign='middle' background="Images/title.gif"> 
        <b><font color="#FFFFFF">�� �� �� �� �� ѹ ��</font></b></td>
  </tr>
  <tr > 
      <td height='150' align='center' valign='middle' bgcolor="#FFF3E6"> <br> <br>
      ѹ��ǰ�������ȱ������ݿ⣬���ⷢ��������� 
      <input name="dbpath" type="hidden" id="dbpath" value="<%=AdminDbPath &mdb%>"> 
      <br> <br> <input type="checkbox" name="boolIs97" value="True">
      �����ʹ�� Access 97 ���ݿ���ѡ��<br>
      (ϵͳĬ��Ϊ Access 2000 ���ݿ�)<br>
        <br> <input name='Submit3' type=Submit value=' ѹ�����ݿ� '>
        <br>
      <br> </td>
  </tr>
</table>
</form>
<%
End Sub


'==========
'����
'==========
Sub BackUpDataYes()

		Dbpath	=	Request.form("Dbpath")
		Dbpath=	server.mappath(Dbpath)
		bkfolder=	Request.form("bkfolder")
		bkdbname=Request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		If fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname & ".asa"
			Else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname & ".asa"
			End If
		Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ ��"&bkfolder&"\"& bkdbname & ".asa��</li>&Action=Yes&ComeUrl="&ComeUrl&""
		Else
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����ʧ�ܣ��Ҳ�����ѡ���Դ���ݿ��ļ����������ݿ������ļ���</li>&Action=OtherErr"
		End If
End Sub

'************
'*��ԭ����    *
'************

Sub RestoreDataYes()		
			dim dbpath,bkfolder,bkdbname,fso,fso1
			Dbpath=Request.form("Dbpath")
			If dbpath="" then
			response.write "<b>�����㱸�ݵ����ݿ�����ƣ������ʼ�ָ����ݿ⣡</b>"	
			Else
			Dbpath=server.mappath(Dbpath&".asa")
			End If
			backpath=server.mappath(Replace(AdminDbPath,"/","\") &Replace(mdb,"/","\") )
		
			Set Fso=server.createobject("scripting.filesystemobject")
			If fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>���ݿ�ָ��ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
Else
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����Ŀ¼��������ѡ����ļ������������룡��</li>&Action=OtherErr"
End If
End Sub
			
'*************
'*���ݿ�ѹ��  *
'*************	
Sub CompactDataYes()

Dim dbpath,boolIs97
dbpath = Request("dbpath")
boolIs97 = Request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If
Const JET_3X = 4
End Sub

Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing

Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>���ݿ�ѹ���ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
Else
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>���ݿ����ƻ�·������ȷ��������ѡ��</li>&Action=OtherErr"
End If
End Function
%>
</body>
</html><br>
<div align="center" class="smalltxt"><%Call MayVote_CopyRight()
Call CloseConn()
%></div>
