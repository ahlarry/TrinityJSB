<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<!-- #include file="Include/MD5.asp"-->
<%
'��Դ��֤
Call CheckUrl()
'����Ա��֤
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
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Բ�����Ч�Ĳ�����"&Action&"��</li>&Action=OtherErr"
End Select

'����û�
Sub AddUsers()
'��ֹ�ǳ�������Ա����
Call CheckUnAdmin1()
Dim UserName,PassWord,System
UserName = ReplaceBadChar(MayHTMLEncode(Trim(Request.Form("UserName"))))
PassWord = Trim(Request.Form("PassWord"))
System = ReplaceBadChar(MayHTMLEncode(Trim(Request.Form("System"))))

If UserName = "" Or Len(UserName) < 4 Or Len(UserName) >20 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û���Ϊ�գ����䳤��С��4������20�ֽڡ�</li>&Action=OtherErr"
If PassWord = "" Or Len(PassWord) < 4 Or Len(PassWord) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�����Ϊ�գ����䳤��С��4������16�ֽڡ�</li>&Action=OtherErr"
If System = "" Or Len(System) >2 Or isInteger(System) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>Ȩ�����ó��������ȷҳ���ύ��</li>&Action=OtherErr"

PassWord = md5(PassWord,16)
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select * From May_Users Where UserName ='"&UserName&"'"
rs.Open SQL,Conn,1,3
If Not(rs.Eof And rs.Bof) Then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����������û����Ѿ����ڣ��뷵���������롣</li>&Action=OtherErr"
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
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û���ӳɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'�����û�
Sub IsLock()
'��ֹ�ǳ�������Ա����
Call CheckUnAdmin1()
Dim UID
UID = Cint(Request("UID"))
If UID = "" Or isInteger(UID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
If UID = Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>������������ʹ�õ��û�������������������Ա��½������������</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.RecordSet")
SQL = "Select UID,IsLock From May_Users Where UID="&UID
rs.Open SQL,conn,1,3
If rs.EOF And rs.BOF Then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�UID�����ڡ�</li>&Action=OtherErr"
Else
	If rs("IsLock") = May_True Then
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>���û��ѱ�������</li>&Action=OtherErr"
	Else
		rs("IsLock") = May_True
		rs.Update
	End If
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û������ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'�������
Sub IsUnLock()
'��ֹ�ǳ�������Ա����
Call CheckUnAdmin1()
Dim UID
UID = Cint(Request("UID"))
If UID = "" Or isInteger(UID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
If UID = Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��ǰ�û�����ʹ��,δ��������</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.RecordSet")
SQL = "Select UID,IsLock From May_Users Where UID="&UID
rs.Open SQL,conn,1,3
If rs.EOF And rs.BOF Then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�UID�����ڡ�</li>&Action=OtherErr"
Else
	If rs("IsLock") = May_False Then
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>���û�û�б�������</li>&Action=OtherErr"
	Else
		rs("IsLock") = May_False
		rs.Update
	End If
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û������ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'ɾ���û�
Sub DelUsers()
'��ֹ�ǳ�������Ա����
Call CheckUnAdmin1()
Dim UID
UID = Cint(Trim(Request.QueryString("UID")))
If UID = "" Or isInteger(UID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
If UID = Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����ɾ������ʹ�õ��û�������������������Ա��½����ɾ����</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From May_Users Where UID="&UID)
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�ɾ���ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'ɾ����ѡ�û�
Sub DelUsersAll()
'��ֹ�ǳ�������Ա����
Call CheckUnAdmin1()
Dim UID
UID = ReplaceBadChar(Trim(Request.Form("UID")))
If UID = "" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From May_Users Where UID In("&UID&")")
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�ɾ���ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'�޸�����
Sub EditPassWord()
Dim UID,OldPassWord,NewPassWord,NewPassWord2
UID = Cint(Trim(Request.Form("UID")))
OldPassWord = Trim(Request.Form("OldPassWord"))
NewPassWord = Trim(Request.Form("NewPassWord"))
NewPassWord2 = Trim(Request.Form("NewPassWord2"))
If UID = "" Or isInteger(UID) = False Or UID <> Session("UID") Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
If OldPassWord = Empty Or Len(OldPassWord) <4 Or Len(OldPassWord) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�������Ϊ�գ����䳤��С��4 �����16�ֽڡ�</li>&Action=OtherErr"
If NewPassWord = Empty Or Len(NewPassWord) <4 Or Len(NewPassWord) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�������Ϊ�գ����䳤��С��4 �����16�ֽڡ�</li>&Action=OtherErr"
If NewPassWord2 = Empty Or Len(NewPassWord2) <4 Or Len(NewPassWord2) > 16 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�ȷ������Ϊ�գ����䳤��С��4 �����16�ֽ�</li>&Action=OtherErr"
If md5(NewPassWord,16) <> md5(NewPassWord2,16) Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û���������ȷ�����벻��ȡ�</li>&Action=OtherErr"

Set rs = Server.Createobject("adodb.RecordSet")
SQL = "Select UID,PassWord From May_Users Where UID="&UID
rs.Open SQL,Conn,1,3
If rs.EOF And rs.BOF Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û������ڡ�</li>&Action=OtherErr"
Else
	If rs("PassWord") <> md5(OldPassWord,16) Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û����������</li>&Action=OtherErr"
	rs("PassWord") = md5(NewPassWord,16)
	rs.Update
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>���������޸ĳɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub
Call CloseConn()
%>