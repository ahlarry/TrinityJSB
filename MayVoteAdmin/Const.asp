<%
Dim AdminDbpath
AdminDbPath = "../"

'==============================
'���ƣ����������֤
'���ã� Call Admin_Setup()
'==============================
Sub CheckUnAdmin()
	If Session("UserName") ="" Or Session("System") ="" Or Session("UID") ="" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Բ�����Ч��ԽȨ������</li>&Action=LoginErr"
End Sub

'��ֹ�ǳ�������Ա����
Sub CheckUnAdmin1()
If Session("System") = 0 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Բ���ԽȨ������</li>&Action=OtherErr"
End Sub
'==============================

'==============================
'���ƣ���Դ��֤
'==============================
Sub CheckUrl()
	From_url = Cstr(Request.ServerVariables("HTTP_REFERER")) 
	Serv_url = Cstr(Request.ServerVariables("SERVER_NAME")) 
	If Mid(From_url,8,Len(Serv_url)) <> Serv_url Then 
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��Դ����</li>&Action=OtherErr"
	End If
End Sub
'==============================
%>