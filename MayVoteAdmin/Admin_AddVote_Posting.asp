<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
'��Դ��֤
Call CheckUrl()
'����Ա��֤
Call CheckUnAdmin()

Action = Request.QueryString("Action")
If Action = "AddVote" Then
	Call AddVote()
ElseIf Action = "AddOption" Then
	Call AddOption()
Else
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Բ�����Ч�Ĳ�����"&Action&"��</li>&Action=OtherErr"
End If

'���ͶƱ��Ŀ
Sub AddVote()
Dim MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_EndDateTime,MayVote_EndDateTime1,MayVote_EndDateTime2,MayVote_EndDateTime3
MayVote_Topic = MayHTMLEncode(Trim(Request.Form("MayVote_Topic")))
MayVote_Check = Trim(Request.Form("MayVote_Check"))
MayVote_Time = Trim(Request.Form("MayVote_Time"))
MayVote_EndDateTime1 = Trim(Request.Form("MayVote_EndDateTime1"))
MayVote_EndDateTime2 = Trim(Request.Form("MayVote_EndDateTime2"))
MayVote_EndDateTime3 = Trim(Request.Form("MayVote_EndDateTime3"))

If MayVote_Topic = "" Or Len(MayVote_Topic) > 50 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ��Ŀ����Ϊ�գ����䳤�ȴ���50�ֽڡ�</li>&Action=OtherErr"
If MayVote_Check = "" Or isInteger(MayVote_Check) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��/�� ѡ�����</li>&Action=OtherErr"
If MayVote_Check <> 1 Then
MayVote_Check = 0
Else
MayVote_Check = 1
End If
If MayVote_Time = "" Or isInteger(MayVote_Time) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ���ʱ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
If MayVote_EndDateTime1 = "" Or Len(MayVote_EndDateTime1) <> 4 Or isInteger(MayVote_EndDateTime1) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ����ʱ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
If MayVote_EndDateTime2 = "" Or Len(MayVote_EndDateTime2) <> 2 Or isInteger(MayVote_EndDateTime2) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ����ʱ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
If MayVote_EndDateTime3 = "" Or Len(MayVote_EndDateTime3) <> 2 Or isInteger(MayVote_EndDateTime3) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ����ʱ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
MayVote_EndDateTime = Cdate(MayVote_EndDateTime1&"-"&MayVote_EndDateTime2&"-"&MayVote_EndDateTime3)

'Set rs = Conn.Execute("Insert Into MayVote_Votes (MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_DateTime,MayVote_EndDateTime) Values ('"&MayVote_Topic&"','"&MayVote_Check&"','"&MayVote_Time&"','#"&NowTime&"#','"&MayVote_EndDateTime&"')")
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select * From MayVote_Votes"
rs.Open SQL,Conn,1,3
rs.AddNew
rs("MayVote_Topic") = MayVote_Topic
rs("MayVote_Check") = MayVote_Check
rs("MayVote_Time") = MayVote_Time
rs("MayVote_DateTime") = Now()
rs("MayVote_EndDateTime") = MayVote_EndDateTime
rs.Update
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ��Ŀ��ӳɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'���ͶƱѡ��
Sub AddOption()
Dim MayVote_Option,MayVote_Nums,MayVote_VotesSelect
MayVote_Option = MayHTMLEncode(Trim(Request.Form("MayVote_Option")))
MayVote_Nums = Trim(Request.Form("MayVote_Nums"))
VID = Trim(Request.Form("VID"))

If MayVote_Option = "" Or Len(MayVote_Option) >20 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱѡ��Ϊ�գ����䳤�ȴ���20�ֽڡ�</li>&Action=OtherErr"
If MayVote_Nums = "" Or isInteger(MayVote_Nums) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>Ĭ��Ʊ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
If VID = "" Or isInteger(VID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����ͶƱ��Ŀδѡ�񣬻���ID��Ϊ������</li>&Action=OtherErr"

Set rs = Server.Createobject("adodb.Recordset")
SQL="Select MayVote_OptionNums From MayVote_Votes Where ID="&VID
rs.Open SQL,Conn,1,3
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>������Ŀ�����ڡ�</li>&Action=OtherErr"
Else
	rs("MayVote_OptionNums") = rs("MayVote_OptionNums") + 1
	MayVote_Order = rs("MayVote_OptionNums")
	rs.Update
End If
rs.Close
Set rs = Nothing
Set rs = Conn.Execute("Insert Into MayVote_Options (MayVote_Option,VID,MayVote_Nums,MayVote_Order) Values ('"&MayVote_Option&"','"&VID&"','"&MayVote_Nums&"','"&MayVote_Order&"')") 
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱѡ����ӳɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

Call CloseConn()
%>