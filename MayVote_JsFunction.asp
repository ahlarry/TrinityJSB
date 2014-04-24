<%
'-------------------JS调用函数---------------------
Function JsForm(ID)
	Set rs = Server.CreateObject("adodb.recordset")
	Sql="Select MayVote_Votes.MayVote_Topic,MayVote_Votes.MayVote_Check,MayVote_Votes.MayVote_Time,MayVote_Options.ID,MayVote_Options.MayVote_Option,MayVote_Options.VID From MayVote_Votes,MayVote_Options Where MayVote_Votes.ID = MayVote_Options.VID And MayVote_Votes.ID="&ID&" Order By MayVote_Options.MayVote_Order Asc"
	rs.Open Sql,Conn,1,1
	If rs.Bof And rs.Eof Then 
		JsForm = "document.write('<div align=center><font color=red>调用发生不知名错误</font></div>')"
	Else
	JsForm = "document.write('<table width=100% border=0 cellspacing=1 cellpadding=2><form name=form method=post action=MayVote_Update.asp?ID="&ID&"><tr><td>　　"&MayHTMLEncode(rs("MayVote_Topic"))&"<input name=MayVote_Time type=hidden id=MayVote_Time value="&rs("MayVote_Time")&"><input name=VID type=hidden id=VID value="&rs("VID")&"></td></tr>');"
	Do While Not rs.Eof
		If rs("MayVote_Check") = May_True Then
			Input = "<input name=ID type=checkbox id=ID value="&rs("ID")&">"
		Else
			Input = "<input type=radio name=ID value="&rs("ID")&">"
		End If
	JsForm =  JsForm + "document.write('<tr><td>　"&Input&"　"&MayHTMLEncode(rs("MayVote_Option"))&"</td></tr>');"
	rs.MoveNext
    Loop
	JsForm = JsForm + "document.write('<tr><td align=center height=50><input name=Submit type=image style=border=0 name=imageField src=Images/MayVote/voteSubmit.gif>&nbsp;&nbsp;<a href=MayVote.asp?Action=Show&ID="&ID&" target=_blank><img src=Images/MayVote/voteView.gif width=52 height=18 border=0></a></td></tr></form></table>');"
	End If
	rs.Close
	Set rs = Nothing
End Function
%>