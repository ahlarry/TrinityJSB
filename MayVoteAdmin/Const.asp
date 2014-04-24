<%
Dim AdminDbpath
AdminDbPath = "../"

'==============================
'名称：管理身份验证
'调用： Call Admin_Setup()
'==============================
Sub CheckUnAdmin()
	If Session("UserName") ="" Or Session("System") ="" Or Session("UID") ="" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>对不起！无效的越权操作。</li>&Action=LoginErr"
End Sub

'禁止非超级管理员访问
Sub CheckUnAdmin1()
If Session("System") = 0 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>对不起，越权操作。</li>&Action=OtherErr"
End Sub
'==============================

'==============================
'名称：来源验证
'==============================
Sub CheckUrl()
	From_url = Cstr(Request.ServerVariables("HTTP_REFERER")) 
	Serv_url = Cstr(Request.ServerVariables("SERVER_NAME")) 
	If Mid(From_url,8,Len(Serv_url)) <> Serv_url Then 
		Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>来源错误。</li>&Action=OtherErr"
	End If
End Sub
'==============================
%>