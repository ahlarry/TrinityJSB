<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%'管理员验证
Call CheckUnAdmin()
'禁止非超级管理员访问
Call CheckUnAdmin1()
%>
<html>
<head>
<title>数据库管理</title>
<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#F6F6F6" leftmargin='2' topmargin='0' marginwidth='0' marginheight='0'>
<br>
<table width='99%' border='1' align='center' cellpadding='3' cellspacing='0' bordercolor="#666666" >
  <tr> 
    <td height='25'  align='center'  background="Images/title.gif"><strong><font color="#FFFFFF">数 
      据 库 管 理</font></strong></td>
  </tr>
  <tr > 
    <td height='30' bgcolor="#FFF3E6" >&nbsp;&nbsp;<strong>管理导航：</strong> <a href='Admin_Database.asp?Action=BackUpData' target=main>备份数据库</a>&nbsp;|&nbsp;<a href='Admin_Database.asp?Action=RestoreData' target=main>恢复数据库</a>&nbsp;|&nbsp;<a href='Admin_Database.asp?Action=CompactData' target=main>压缩数据库</a></td>
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
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法参数：Action=<font color=red>"&Action&"</font></li>&Action=OtherErr"
		End If
'************
'*  备份选择  *
'************
		Sub BackUpData()
%>
<form action='Admin_Database.asp?Action=BackUpDataYes' method='post' name="backup" id="backup" target="main">
  <table width='99%' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#666666" >
    <tr>
      <td height='25' align='center' valign='middle' background="Images/title.gif"> 
        <b><font color="#FFFFFF">备 份 数 据 库</font></b> </td> 
</tr>
<tr >
      <td height='150' align='center' valign='middle' bgcolor="#FFF3E6"> 
        <table width='99%' border='1' cellpadding='3' cellspacing='0' bordercolor="#FFFFFF">
          <tr> 
            <td width=30% height='34' align='right'>当前数据库路径：</td>
            <td width=30%> 
              <input type="hidden" size="20" name="DBpath" value="<%=Replace(AdminDbPath,"/","\") &Replace(mdb,"/","\") %>"> 
              <%=AdminDbPath&mdb%></td>
            <td width=40%>数据库所在目录，相对路径</td>
          </tr>
          <tr> 
            <td height='34' align='right'>备份目录：</td>
            <td><input type="text" size="20" name="bkfolder" value="Databackup" readonly> </td>
            <td>相对路径目录，如目录不存在，将自动创建<br>默认路径，最好不要更改</td>
          </tr>
          <tr> 
            <td height='34' align='right'>备份名称：</td>
            <td height='34'><input type="text" size="20" name="bkDBname" value='MayVote_backup'> 
            </td>
            <td height='34'>如有同名文件，将覆盖<br>
              <font color="#FF0000">文件后缀强制为 .asa ,请不要再填写文件后缀</font></td>
          </tr>
          <tr> 
            <td height='20' colspan='3' align="center"> 
              <input name='Submit' type="Submit" value=' 开始备份 '> 
              <input name="Action" type="hidden" id="Action" value="BackUpYes"></td>
          </tr>
          <tr> 
            <td height='20' colspan='3'>( 备份数据库需要FSO支持，FSO相关帮助请看微软网站 ) <br>
              -----------------------------------------------------------------------------------------<br>
              在上面填写本程序的数据库路径全名，本程序的默认数据库文件为../Data/MayVote.mdb，请一定不能用默认名称命名备份数据库<br>
              您可以用这个功能来备份您的论坛数据，以保证您的数据安全！<br>
              注意：所有路径都是相对与程序空间根目录的相对路径 </td>
          </tr>
        </table>
</td>
</tr>
</table>
</form>
<%End Sub

'************
'*还原选择    *
'************
Sub RestoreData()
%>
<form action="Admin_Database.asp?Action=RestoreDataYes" method='post' name="RestoreYes" id="RestoreYes" target="main">
  <table width='99%' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#666666" >
    <tr class='title'> 
      <td height='25' align='center' valign='middle' background="Images/title.gif"><b><font color="#FFFFFF">数 
        据 库 恢 复</font></b></td>
  </tr>
    <tr > 
      <td bgcolor="#FFF3E6"> 
        <table width='99%' border='1' cellpadding='3' cellspacing='0' bordercolor="#FFFFFF">
          <tr> 
            <td width='30%' height='30' align='right'>备份数据库路径（相对）：</td>
            <td width="70%" height='30' align="left"> 
              <input name="dbpath" type=text id="dbpath" value="Databackup\MayVote_Backup" size="50" maxlength="200">
              <strong><font color="#FF0000"> </font></strong></td>
          </tr>
          <tr> 
            <td height='20' colspan='2' align="center"> 
              <input name='Submit' type="Submit" value=" 恢复数据 "> 
              <input name="Action" type="hidden" id="Action" value="RestoreYes"> 
            </td>
          </tr>
          <tr> 
            <td height='20' colspan='2'>填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确），然后修改Conn.asp文件，如果目标文件名和当前使用数据库名一致的话，不需修改Conn.asp文件<br>
              -----------------------------------------------------------------------------------------<br>
              在上面填写本程序的数据库路径全名，本程序的默认备份数据库文件为DataBackup/MayVote_Backup.asa，请按照您的备份文件自行修改。<br>
              您可以用这个功能来恢复您的论坛数据，以保证您的数据安全！<br>
              注意：1.所有路径都是相对与程序空间根目录的相对路径 <br>
              　　　2.<font color="#FF0000">数据库备份文件后缀已经强制为.asa</font> </td>
          </tr>
        </table>
        
      </td>
  </tr>
</table>
</form>
<%
End Sub

'**************
'*数据库库压缩*
'**************
Sub CompactData()
%>
<form action="Admin_Database.asp?Action=CompactDataYes" method='post' name="Compact" id="Compact" target="main">
  <table width='99%' border='1' align='center' cellpadding='0' cellspacing='0' bordercolor="#666666" >
    <tr class='title'> 
      <td height='25' align='center' valign='middle' background="Images/title.gif"> 
        <b><font color="#FFFFFF">数 据 库 在 线 压 缩</font></b></td>
  </tr>
  <tr > 
      <td height='150' align='center' valign='middle' bgcolor="#FFF3E6"> <br> <br>
      压缩前，建议先备份数据库，以免发生意外错误。 
      <input name="dbpath" type="hidden" id="dbpath" value="<%=AdminDbPath &mdb%>"> 
      <br> <br> <input type="checkbox" name="boolIs97" value="True">
      如果您使用 Access 97 数据库请选择<br>
      (系统默认为 Access 2000 数据库)<br>
        <br> <input name='Submit3' type=Submit value=' 压缩数据库 '>
        <br>
      <br> </td>
  </tr>
</table>
</form>
<%
End Sub


'==========
'备份
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
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>备份数据库成功！您备份的数据库路径为 ："&bkfolder&"\"& bkdbname & ".asa。</li>&Action=Yes&ComeUrl="&ComeUrl&""
		Else
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>备份失败，找不到您选择的源数据库文件，请检查数据库连接文件。</li>&Action=OtherErr"
		End If
End Sub

'************
'*还原程序    *
'************

Sub RestoreDataYes()		
			dim dbpath,bkfolder,bkdbname,fso,fso1
			Dbpath=Request.form("Dbpath")
			If dbpath="" then
			response.write "<b>输入你备份的数据库的名称，点击开始恢复数据库！</b>"	
			Else
			Dbpath=server.mappath(Dbpath&".asa")
			End If
			backpath=server.mappath(Replace(AdminDbPath,"/","\") &Replace(mdb,"/","\") )
		
			Set Fso=server.createobject("scripting.filesystemobject")
			If fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>数据库恢复成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
Else
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>备份目录下无您所选择的文件，请重新输入！。</li>&Action=OtherErr"
End If
End Sub
			
'*************
'*数据库压缩  *
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
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>数据库压缩成功！</li>&Action=Yes&ComeUrl="&ComeUrl&""
Else
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>数据库名称或路径不正确，请重新选择！</li>&Action=OtherErr"
End If
End Function
%>
</body>
</html><br>
<div align="center" class="smalltxt"><%Call MayVote_CopyRight()
Call CloseConn()
%></div>
