<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
'来源验证
Call CheckUrl()
'管理员验证
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
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>对不起！无效的参数："&Action&"。</li>&Action=OtherErr"
End Select

'编辑投票项目
Sub EditVote()
Dim ID,MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_EndDateTime,MayVote_EndDateTime1,MayVote_EndDateTime2,MayVote_EndDateTime3
ID = Trim(Request.Form("ID"))
MayVote_Topic = MayHTMLEncode(Trim(Request.Form("MayVote_Topic")))
MayVote_Check = Trim(Request.Form("MayVote_Check"))
MayVote_Time = Trim(Request.Form("MayVote_Time"))
MayVote_EndDateTime1 = Trim(Request.Form("MayVote_EndDateTime1"))
MayVote_EndDateTime2 = Trim(Request.Form("MayVote_EndDateTime2"))
MayVote_EndDateTime3 = Trim(Request.Form("MayVote_EndDateTime3"))

If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的ID参数。</li>&Action=OtherErr"
If MayVote_Topic = "" Or Len(MayVote_Topic) >50 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票项目标题为空，或其长度大于50字节。</li>&Action=OtherErr"
If MayVote_Check = "" Or isInteger(MayVote_Check) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>单/多 选属性错误。</li>&Action=OtherErr"
If MayVote_Check <> 1 Then
MayVote_Check = 0
Else
MayVote_Check = 1
End If
If MayVote_Time = "" Or isInteger(MayVote_Time) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票间隔时间为空，或其ID不为正整。</li>&Action=OtherErr"
If MayVote_EndDateTime1 = "" Or Len(MayVote_EndDateTime1) <> 4 Or isInteger(MayVote_EndDateTime1) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票结束时间为空，或其不为正整数。</li>&Action=OtherErr"
If MayVote_EndDateTime2 = "" Or Len(MayVote_EndDateTime2) <> 2 Or isInteger(MayVote_EndDateTime2) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票结束时间为空，或其不为正整数。</li>&Action=OtherErr"
If MayVote_EndDateTime3 = "" Or Len(MayVote_EndDateTime3) <> 2 Or isInteger(MayVote_EndDateTime3) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票结束时间为空，或其不为正整数。</li>&Action=OtherErr"
MayVote_EndDateTime = Cdate(MayVote_EndDateTime1&"-"&MayVote_EndDateTime2&"-"&MayVote_EndDateTime3)

Set rs = Server.Createobject("adodb.Recordset")
SQL="Select MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_EndDateTime From MayVote_Votes Where ID="&ID
rs.Open SQL,Conn,1,3
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>你所修改的数据不存在。</li>&Action=OtherErr"
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
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票项目更新成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'删除投票项目
Sub DelVote()
Dim ID
ID = Trim(Request.QueryString("ID"))
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From MayVote_Votes Where ID="&ID)
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票项目删除成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'编辑投票选项
Sub EditOption()
Dim ID,MayVote_Option,MayVote_Nums,VID,MayVote_Order
ID = Trim(Request.Form("ID"))
MayVote_Option = MayHTMLEncode(Trim(Request.Form("MayVote_Option")))
MayVote_Nums = Trim(Request.Form("MayVote_Nums"))
VID = Trim(Request.Form("VID"))
MayVote_Order = Trim(Request.Form("MayVote_Order"))

If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
If MayVote_Option = "" Or Len(MayVote_Option) >20 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票选项为空，或其长度大于20字节。</li>&Action=OtherErr"
If MayVote_Nums = "" Or isInteger(MayVote_Nums) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>默认票数为空，或其不为正整数。</li>&Action=OtherErr"
If VID = "" Or isInteger(VID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>所属投票项目未选择，或其ID不为正整数。</li>&Action=OtherErr"
If MayVote_Order = "" Or isInteger(MayVote_Order) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>显示顺序为空，或其ID不为正整。</li>&Action=OtherErr"

Set rs = Server.Createobject("adodb.Recordset")
SQL="Select MayVote_Option,MayVote_Nums,VID,MayVote_Order From MayVote_Options Where ID="&ID
rs.Open SQL,Conn,1,3
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>你所修改的数据不存在。</li>&Action=OtherErr"
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
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票选项更新成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'删除投票选项
Sub DelOption()
Dim ID
ID = Trim(Request.QueryString("ID"))
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的UID参数。</li>&Action=OtherErr"
Set rs = Conn.Execute("Delete * From MayVote_Options Where ID="&ID)
Set rs = Nothing
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票选项删除成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'更新所有
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
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>投票显示顺序更新成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

Call CloseConn()
%>