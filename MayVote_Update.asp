<!-- #include file="MayVote_Conn.asp"-->
<!-- #include file="MayVoteAdmin/Include/MayVote_Function.asp"-->
<%'来源验证
From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
Dim userip,ip1,ip2
ip1=request.servervariables("http_x_forwarded_for")
ip2=request.servervariables("remote_addr")
if instr(ip1,",")>0 then ip1=left(ip1,instr(ip1,",")-1)
if instr(ip2,",")>0 then ip2=left(ip2,instr(ip2,",")-1)
if ip1 <> "" then
		userip = ip1
else
		userip = ip2
end if
'If Mid(From_url,8,Len(Serv_url)) <> Serv_url Then
'	Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>禁止异地提交数据。</li>&Action=OtherErr"
'End If
'结束
Dim ID,MayVote_Time
ID = ReplaceBadChar(Trim(Request.Form("ID")))
VID = Trim(Request.Form("VID"))
MayVote_Time = Trim(Request.Form("MayVote_Time"))
If ID = "" Then Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>非法的ID参数。</li>&Action=OtherErr"
If VID = "" Or isInteger(VID) = False Then Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>非法的ID参数。</li>&Action=OtherErr"
If MayVote_Time = "" Or isInteger(MayVote_Time) = False Then Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>非法的时间参数。</li>&Action=OtherErr"
'COOKIES保存时间设置
If MayVote_Time = 0 Then
MayVote_Time = ""
Else
	If MayVote_Time >10 Then
		MayVote_Time = MayVote_Time/10
	End if
MayVote_Time = NowTime + MayVote_Time
End If
If Request.Cookies("MayVote"&VID&"") <> Empty Then Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>对不起！您已经投过票了。</li>&Action=OtherErr"
Set rs1 = Conn.Execute("Select VotedIP From MayVote_Ed Where Vid ='"&VID&"'")
do while not rs1.Eof
	If userip = rs1("VotedIP") Then
		Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>对不起,您已经投过票了！</li>&Action=OtherErr"
		Exit Do
	End If
	rs1.movenext
loop
rs1.close
Set rs1 = Nothing

'Set rs = Server.CreateObject("Adodb.Recordset")
'Sql="Select VotedIP From MayVote_Ed Where Vid ='"&VID&"'"
'rs.Open Sql,Conn,1,1
'do while not rs.Eof Then
'	If userip = rs("VotedIP") Then
'		Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>对不起,您已经投过票了！</li>&Action=OtherErr"
'		Exit Do
'	End If
'	rs.movenext
'loop
'rs.close

Set rs = Server.CreateObject("Adodb.Recordset")
Sql="Select ID From MayVote_Options Where ID In("&ID&")"
rs.Open Sql,Conn,1,1
If rs.Eof And rs.Bof Then
Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>对不起，没有这个投票项目。</li>&Action=OtherErr"
Else
Set rs2 = Conn.Execute("Select MayVote_EndDateTime From MayVote_Votes Where ID ="&VID&"")
If Now() >= rs2("MayVote_EndDateTime") Then Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>对不起，这个投票项目已经过期。</li>&Action=OtherErr"
Set rs2 = Nothing
Set rs1 = Conn.Execute("Update MayVote_Options Set MayVote_Nums = MayVote_Nums+1 Where ID In("&ID&")")
Set rs1 = Nothing
Set rs2 = Conn.Execute("insert into [MayVote_Ed] (Vid, VotedIP) values ('"&VID&"','"&userip&"')")
Set rs2 = Nothing
Response.Cookies("MayVote"&VID&"") = "Yes"
Response.Cookies("MayVote"&VID&"").Expires = MayVote_Time
Dim ComeUrl
ComeUrl = "MayVote.asp?ID="&VID&""
Response.Redirect"MayVote_Info.asp?InfoContent=<li>操作成功！</li>&Action=Yes&ComeUrl="&ComeUrl&""
End If
rs.Close
Set rs = Nothing
Call CloseConn()
%>