<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
'��Դ��֤
Call CheckUrl()
'����Ա��֤
Call CheckUnAdmin()
Action = Request.QueryString("Action")
Select Case Action
	Case "EditVote"
		Call EditVote()
	Case "DelVote"
		Call DelVote()
	Case "EditOption"
		Call EditOption()
	Case "DelOption"
		Call DelOption()
	Case "AllUpdate" 
		Call AllUpdate()
	Case Else
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Բ�����Ч�Ĳ�����"&Action&"��</li>&Action=OtherErr"
End Select

'�༭ͶƱ��Ŀ
Sub EditVote()
Dim ID,MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_EndDateTime,MayVote_EndDateTime1,MayVote_EndDateTime2,MayVote_EndDateTime3
ID = Trim(Request.Form("ID"))
MayVote_Topic = MayHTMLEncode(Trim(Request.Form("MayVote_Topic")))
MayVote_Check = Trim(Request.Form("MayVote_Check"))
MayVote_Time = Trim(Request.Form("MayVote_Time"))
MayVote_EndDateTime1 = Trim(Request.Form("MayVote_EndDateTime1"))
MayVote_EndDateTime2 = Trim(Request.Form("MayVote_EndDateTime2"))
MayVote_EndDateTime3 = Trim(Request.Form("MayVote_EndDateTime3"))

If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���ID������</li>&Action=OtherErr"
If MayVote_Topic = "" Or Len(MayVote_Topic) >50 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ��Ŀ����Ϊ�գ����䳤�ȴ���50�ֽڡ�</li>&Action=OtherErr"
If MayVote_Check = "" Or isInteger(MayVote_Check) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��/�� ѡ���Դ���</li>&Action=OtherErr"
If MayVote_Check <> 1 Then
MayVote_Check = 0
Else
MayVote_Check = 1
End If
If MayVote_Time = "" Or isInteger(MayVote_Time) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ���ʱ��Ϊ�գ�����ID��Ϊ������</li>&Action=OtherErr"
If MayVote_EndDateTime1 = "" Or Len(MayVote_EndDateTime1) <> 4 Or isInteger(MayVote_EndDateTime1) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ����ʱ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
If MayVote_EndDateTime2 = "" Or Len(MayVote_EndDateTime2) <> 2 Or isInteger(MayVote_EndDateTime2) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ����ʱ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
If MayVote_EndDateTime3 = "" Or Len(MayVote_EndDateTime3) <> 2 Or isInteger(MayVote_EndDateTime3) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ����ʱ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
MayVote_EndDateTime = Cdate(MayVote_EndDateTime1&"-"&MayVote_EndDateTime2&"-"&MayVote_EndDateTime3)

Set rs = Server.Createobject("adodb.Recordset")
SQL="Select MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_EndDateTime From MayVote_Votes Where ID="&ID
rs.Open SQL,Conn,1,3
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�����޸ĵ����ݲ����ڡ�</li>&Action=OtherErr"
Else
	rs("MayVote_Topic") = MayVote_Topic
	rs("MayVote_Check") = MayVote_Check
	rs("MayVote_Time") = MayVote_Time
	rs("MayVote_EndDateTime") = MayVote_EndDateTime
	rs.Update
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = "Admin_Edit.asp"
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ��Ŀ���³ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'ɾ��ͶƱ��Ŀ
Sub DelVote()
Dim ID
ID = Trim(Request.QueryString("ID"))
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From MayVote_Votes Where ID="&ID)
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ��Ŀɾ���ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'�༭ͶƱѡ��
Sub EditOption()
Dim ID,MayVote_Option,MayVote_Nums,VID,MayVote_Order
ID = Trim(Request.Form("ID"))
MayVote_Option = MayHTMLEncode(Trim(Request.Form("MayVote_Option")))
MayVote_Nums = Trim(Request.Form("MayVote_Nums"))
VID = Trim(Request.Form("VID"))
MayVote_Order = Trim(Request.Form("MayVote_Order"))

If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
If MayVote_Option = "" Or Len(MayVote_Option) >20 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱѡ��Ϊ�գ����䳤�ȴ���20�ֽڡ�</li>&Action=OtherErr"
If MayVote_Nums = "" Or isInteger(MayVote_Nums) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>Ĭ��Ʊ��Ϊ�գ����䲻Ϊ��������</li>&Action=OtherErr"
If VID = "" Or isInteger(VID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����ͶƱ��Ŀδѡ�񣬻���ID��Ϊ��������</li>&Action=OtherErr"
If MayVote_Order = "" Or isInteger(MayVote_Order) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��ʾ˳��Ϊ�գ�����ID��Ϊ������</li>&Action=OtherErr"

Set rs = Server.Createobject("adodb.Recordset")
SQL="Select MayVote_Option,MayVote_Nums,VID,MayVote_Order From MayVote_Options Where ID="&ID
rs.Open SQL,Conn,1,3
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�����޸ĵ����ݲ����ڡ�</li>&Action=OtherErr"
Else
	rs("MayVote_Option") = MayVote_Option
	rs("MayVote_Nums") = MayVote_Nums
	rs("VID") = VID
	rs("MayVote_Order") = MayVote_Order
	rs.Update
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = "Admin_Edit.asp"
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱѡ����³ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'ɾ��ͶƱѡ��
Sub DelOption()
Dim ID
ID = Trim(Request.QueryString("ID"))
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���UID������</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From MayVote_Options Where ID="&ID)
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱѡ��ɾ���ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'��������
Sub AllUpdate()
Dim OptionID,MayVote_Order
OptionID = Request.Form("OptionID")
OptionID = Split(OptionID,",")
MayVote_Order = Request.Form("MayVote_Order")
MayVote_Order = Split(MayVote_Order,",")

For i = 0 To Ubound(OptionID)
Set rs = Conn.Execute("Update MayVote_Options Set MayVote_Order ='"&Trim(MayVote_Order(i))&"' Where ID="&Trim(OptionID(i))&" ")
Set rs = Nothing
Next
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>ͶƱ��ʾ˳����³ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

Call CloseConn()
%>