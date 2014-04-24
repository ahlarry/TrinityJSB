<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<!-- #include file="Include/MD5.asp"-->
<%
'来源验证
Call CheckUrl()
'管理员验证
Call CheckUnAdmin()

Action = Request("Action")
Select Case Action
	Case "AddUsers"
		Call AddUsers()
	Case "IsLock"
		Call IsLock()
	Case "IsUnLock"
		Call IsUnLock()
	Case "DelUsersAll"
		Call DelUsersAll()
	Case "DelUsers"
		Call DelUsers()
	Case "EditPassWord"
		Call EditPassWord()
	Case Else
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>对不起！无效的参数："&Action&"。</li>&Action=OtherErr"
End Select

'添加用户
Sub AddUsers()
'禁止非超级管理员访问
Call CheckUnAdmin1()
Dim UserName,PassWord,System
UserName = ReplaceBadChar(MayHTMLEncode(Trim(Request.Form("UserName"))))
PassWord = Trim(Request.Form("PassWord"))
System = ReplaceBadChar(MayHTMLEncode(Trim(Request.Form("System"))))

If UserName = "" Or Len(UserName) < 4 Or Len(UserName) >20 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户名为空，或其长度小于4、大于20字节。</li>&Action=OtherErr"
If PassWord = "" Or Len(PassWord) < 4 Or Len(PassWord) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户密码为空，或其长度小于4、大于16字节。</li>&Action=OtherErr"
If System = "" Or Len(System) >2 Or isInteger(System) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>权限设置出错，请从正确页面提交。</li>&Action=OtherErr"

PassWord = md5(PassWord,16)
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select * From May_Users Where UserName ='"&UserName&"'"
rs.Open SQL,Conn,1,3
If Not(rs.Eof And rs.Bof) Then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>您所输入的用户名已经存在，请返回重新输入。</li>&Action=OtherErr"
Else
rs.AddNew
rs("UserName") = UserName
rs("PassWord") = PassWord
rs("System") = System
rs.Update
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户添加成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'锁定用户
Sub IsLock()
'禁止非超级管理员访问
Call CheckUnAdmin1()
Dim UID
UID = Cint(Request("UID"))
If UID = "" Or isInteger(UID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
If UID = Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>不能锁定正在使用的用户，请用其他超级管理员登陆后再行锁定。</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.RecordSet")
SQL = "Select UID,IsLock From May_Users Where UID="&UID
rs.Open SQL,conn,1,3
If rs.EOF And rs.BOF Then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户UID不存在。</li>&Action=OtherErr"
Else
	If rs("IsLock") = May_True Then
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>该用户已被锁定。</li>&Action=OtherErr"
	Else
		rs("IsLock") = May_True
		rs.Update
	End If
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户锁定成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'解除锁定
Sub IsUnLock()
'禁止非超级管理员访问
Call CheckUnAdmin1()
Dim UID
UID = Cint(Request("UID"))
If UID = "" Or isInteger(UID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
If UID = Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>当前用户正在使用,未被加锁。</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.RecordSet")
SQL = "Select UID,IsLock From May_Users Where UID="&UID
rs.Open SQL,conn,1,3
If rs.EOF And rs.BOF Then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户UID不存在。</li>&Action=OtherErr"
Else
	If rs("IsLock") = May_False Then
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>该用户没有被锁定。</li>&Action=OtherErr"
	Else
		rs("IsLock") = May_False
		rs.Update
	End If
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户解锁成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'删除用户
Sub DelUsers()
'禁止非超级管理员访问
Call CheckUnAdmin1()
Dim UID
UID = Cint(Trim(Request.QueryString("UID")))
If UID = "" Or isInteger(UID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
If UID = Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>不能删除正在使用的用户，请用其他超级管理员登陆后再删除。</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From May_Users Where UID="&UID)
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户删除成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'删除多选用户
Sub DelUsersAll()
'禁止非超级管理员访问
Call CheckUnAdmin1()
Dim UID
UID = ReplaceBadChar(Trim(Request.Form("UID")))
If UID = "" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From May_Users Where UID In("&UID&")")
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户删除成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'修改密码
Sub EditPassWord()
Dim UID,OldPassWord,NewPassWord,NewPassWord2
UID = Cint(Trim(Request.Form("UID")))
OldPassWord = Trim(Request.Form("OldPassWord"))
NewPassWord = Trim(Request.Form("NewPassWord"))
NewPassWord2 = Trim(Request.Form("NewPassWord2"))
If UID = "" Or isInteger(UID) = False Or UID <> Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
If OldPassWord = Empty Or Len(OldPassWord) <4 Or Len(OldPassWord) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户旧密码为空，或其长度小于4 或大于16字节。</li>&Action=OtherErr"
If NewPassWord = Empty Or Len(NewPassWord) <4 Or Len(NewPassWord) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户新密码为空，或其长度小于4 或大于16字节。</li>&Action=OtherErr"
If NewPassWord2 = Empty Or Len(NewPassWord2) <4 Or Len(NewPassWord2) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户确认密码为空，或其长度小于4 或大于16字节</li>&Action=OtherErr"
If md5(NewPassWord,16) <> md5(NewPassWord2,16) Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户新密码与确认密码不相等。</li>&Action=OtherErr"

Set rs = Server.Createobject("adodb.RecordSet")
SQL = "Select UID,PassWord From May_Users Where UID="&UID
rs.Open SQL,Conn,1,3
If rs.EOF And rs.BOF Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户不存在。</li>&Action=OtherErr"
Else
	If rs("PassWord") <> md5(OldPassWord,16) Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户旧密码错误。</li>&Action=OtherErr"
	rs("PassWord") = md5(NewPassWord,16)
	rs.Update
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>个人密码修改成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub
Call CloseConn()
%>