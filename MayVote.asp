<!-- #include file="MayVote_Conn.asp"-->
<!-- #include file="MayVoteAdmin/Include/MayVote_Function.asp"-->
<!-- #include file="MayVote_JsFunction.asp"-->
<%
Action = Request.QueryString("Action")
If Action = "Show" Then
	Call Show()
ElseIf Action = "JS" Then
	Call JsTransfer()
Else
	Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>�Բ�����Ч�Ĳ�����"&Action&"��</li>&Action=OtherErr"
End If

'��ʾ��ϸ��
Sub Show()
ID = Request.QueryString("ID")
If ID = Empty Or isInteger(ID) = False Then Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>�Ƿ���ID������</li>&Action=OtherErr"
Set rs = Server.CreateObject("Adodb.Recordset")
Sql="Select MayVote_Votes.MayVote_Topic,MayVote_Votes.MayVote_DateTime,MayVote_Votes.MayVote_EndDateTime,MayVote_Votes.MayVote_Check,MayVote_Votes.MayVote_Time,MayVote_Options.ID,MayVote_Options.MayVote_Option,MayVote_Options.MayVote_Nums From MayVote_Votes,MayVote_Options Where MayVote_Votes.ID = MayVote_Options.VID And MayVote_Votes.ID="&ID&" Order By MayVote_Options.MayVote_Order Asc"
rs.Open Sql,Conn,1,1
If rs.Eof And rs.Bof Then
Response.Redirect "MayVote_Info.asp?InfoContent=<br><li>�Բ���û�����ͶƱ��Ŀ��</li>&Action=OtherErr"
Else
Dim MayVote_Check
MayVote_Check = rs("MayVote_Check")'�Ƿ��ѡ
'��ֵ����
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% = rs("MayVote_Topic")&" - "&Application("MayVote_Name")%> - MayVote ͶƱϵͳ</title>
<link href="Images/MayVote/Style.css" rel="stylesheet" type="text/css">
</head>
<body>
<table width="600" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
  <tr>
    <td height="24" background="Images/MayVote/BG_Title.jpg">&nbsp;<img src="Images/MayVote/Light.gif" align="absmiddle"> <b>��վ����<font color="red">��<%If Len(rs("MayVote_Topic")) > 20 Then
MayVote_Topic = Left(rs("MayVote_Topic"),27)&"..."
Else
MayVote_Topic = rs("MayVote_Topic")
End If
Response.Write MayHTMLEncode(MayVote_Topic)%>��</font>�ĵ�����<%
If MayVote_Check = May_True Then Response.Write "<font color='red'>(��ѡ)</font>"%></b></td>
  </tr>
  <tr>
    <td><br><div id=list3>��<b>�������⣺</b><% = MayHTMLEncode(rs("MayVote_Topic"))%></div><br><b>��Ŀǰ����Ʊ��Ϊ��</b>
<%
Set rs1 = Server.Createobject("adodb.Recordset")
Sql = "Select Sum(MayVote_Nums) As Nums From MayVote_Options Where VID="&ID
rs1.Open sql,conn,1,1
Response.Write"<font color='red'>"&rs1("Nums")&"Ʊ</font>"
If rs1("Nums") = 0 Then Response.Write"&nbsp;&nbsp;&nbsp;&nbsp;<b>Ŀǰ��û����ͶƱ��</b>"
Response.Write"<br>&nbsp;&nbsp;ͶƱ��ʼʱ�䣺"&FormatDateTime(rs("MayVote_DateTime"),vbLongDate)&"&nbsp;&nbsp;&nbsp;&nbsp;ͶƱ����ʱ�䣺"&FormatDateTime(rs("MayVote_EndDateTime"),vbLongDate)&"<br><br>"
While Not rs.Eof%>
<table width="540" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="1" height="23" valign="middle">&nbsp;&nbsp;</td>
        <td width="539" valign="middle" background="images/MayVote/BG_Options.gif">&nbsp;��ѡ�<b><% = MayHTMLEncode(rs("MayVote_Option")) %></b></td>
      </tr>
      <tr>
        <td colspan="2"><table width="500" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="4"></td>
            </tr>
          </table>
            <table width="400" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="80" height="21" align="right" valign="middle">��Ʊ�ʣ�</td>
                <td width="4" valign="top"><img src="images/MayVote/left.gif" width="4" height="21"></td>
                <td valign="top"><img src="images/MayVote/greenbar.gif" width="<%If rs1("Nums") > 0 Then
Response.Write Int(rs("MayVote_Nums")/rs1("Nums")*100*3)
Else
Response.Write "0"
End If%>" height="21"></td>
                <td width="6" align="center" valign="middle"><img src="images/MayVote/mid.gif" width="6" height="21"> </td>
                <td valign="top"><img src="images/MayVote/whitebar.gif" width="<%If rs1("Nums") > 0 Then
Response.Write Int(305-rs("MayVote_Nums")/rs1("Nums")*100*3)
Else
Response.Write "305"
End If%>" height="21"></td>
                <td width="6" valign="top"><img src="images/MayVote/right.gif" width="6" height="21"></td>
              </tr>
          </table></td>
      </tr>
      <tr>
        <td colspan="2"><table width="530" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="80">&nbsp;</td>
              <td width="<%If rs1("Nums") > 0 Then
Response.Write Int(rs("MayVote_Nums")/rs1("Nums")*100*3)
Else
Response.Write "0"
End If%>">&nbsp;</td>
              <td width="<%If rs1("Nums") > 0 Then
Response.Write Int(450-rs("MayVote_Nums")/rs1("Nums")*100*3)
Else
Response.Write "450"
End If%>">ռ��<%If rs1("Nums") > 0 Then
Response.Write Round(rs("MayVote_Nums")/rs1("Nums")*100,2)
Else
Response.Write "0"
End If%>%[��<font color="#FF0000"><%= rs("MayVote_Nums") %></font>Ʊ]</td>
            </tr>
        </table></td>
      </tr>
    </table>
<%
rs.MoveNext
Wend
End If
rs1.CLose
Set rs1 = Nothing
rs.Close
Set rs = Nothing
%></td>
  </tr>
</table>
<br>
<table width="600" border="0" align="center" cellpadding="0" cellspacing="3">
  <tr>
    <td width="300" valign="top"><%'��֤COOKIES�Ƿ�ͶƱ
Dim MayVote_Already,userip,ip1,ip2
If Request.Cookies("MayVote"&ID&"") <> Empty Then
MayVote_Already = May_True
Else
MayVote_Already = May_False
End If
ip1=request.servervariables("http_x_forwarded_for")
ip2=request.servervariables("remote_addr")
if instr(ip1,",")>0 then ip1=left(ip1,instr(ip1,",")-1)
if instr(ip2,",")>0 then ip2=left(ip2,instr(ip2,",")-1)
if ip1 <> "" then
		userip = ip1
else
		userip = ip2
end if
Set rs2 = Conn.Execute("Select VotedIP From MayVote_Ed Where Vid ='"&ID&"'")
do while not rs2.eof
	If userip = rs2("VotedIP") Then
		MayVote_Already = May_True
	End If
	rs2.movenext
loop
rs2.close
Set rs2 = Nothing
%><table width="300" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
          <td height="22" background="Images/MayVote/BG_Title.jpg">&nbsp;<img src="Images/MayVote/Light.gif" align="absmiddle"> <b><%If MayVote_Already = May_True Then
Response.Write "���Ѿ�Ͷ��Ʊ�ˣ�лл���Ĳ��룡"
Else
Response.Write "����û��ͶƱ�������ڴ�Ͷ���������һƱ��"
End If%></b></td>
        </tr>
          <tr>
            <td><%
If MayVote_Already = May_Flase Then
Response.Write "<script language='JavaScript' src='MayVote.asp?Action=JS&ID="&ID&"'></script>"
Else
Response.Write "<div id=list4><font size='+7' color='Green'>��</font></div>"
End If
%></td>
</tr>
</table>
</td>
<td valign="top"><%Set rsSum = Server.Createobject("Adodb.Recordset")
Sql="Select Count(ID) From MayVote_Votes"
rsSum.Open Sql,Conn,1,1
If rsSum(0) >1 Then
%>
  <table width="300" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
        <tr>
          <td height="22" background="Images/MayVote/BG_Title.jpg">&nbsp;<img src="Images/MayVote/Light.gif" align="absmiddle"> <b>��ӭ�������μӱ�վ����������</b></td>
        </tr>
        <tr>
          <td height="102" valign="top"><%
Dim MayVote_Topic
Set rs = Server.Createobject("Adodb.Recordset")
Sql="Select Top 5 ID,MayVote_Topic From MayVote_Votes Where ID <> "&ID&" And Now() < MayVote_EndDateTime Order By ID Desc"
rs.Open Sql,Conn,1,1
While Not rs.Eof
If Len(rs("MayVote_Topic")) >20 Then
MayVote_Topic = Left(rs("MayVote_Topic"),20)&"��"
Else
MayVote_Topic = rs("MayVote_Topic")
End If
Response.Write "<div id=list1>��<font color='red'>��</font><a href='MayVote.asp?Action=Show&ID="&rs("ID")&"'>"&MayHTMLEncode(MayVote_Topic)&"</a></div>"
rs.MoveNext
Wend
rs.Close
Set rs = Nothing
%></td>
        </tr>
      </table>
<%End If
rsSum.Close
Set rsSum = Nothing%></td>
  </tr>
</table>
<div align="center" class="smalltxt"><br><br><% = Application("MayVote_Copy")%><br><% Call MayVote_CopyRight()%><br>ҳ��ִ��ʱ��: <%Call CountTime()%> ��.</div>
</body>
</html>
<%
End Sub

'���ñ�
Sub JsTransfer()
Dim ID,Rw
ID = Request.QueryString("ID")
If ID = "" Or isInteger(ID) = False Then
	Response.Write "document.write('<div align=center><font color=red>ͶƱID����</font></div>')"
	Exit Sub
End If
'�ж�����(���Ϊ���������κε�ַ,�������ַ�Ƿ��������)
If Application("MayVote_Setting") <> Empty Then
		If instr(Application("MayVote_Setting"),Request.ServerVariables("SERVER_NAME")) < 1 Then
			Response.Write "document.write('<div align=center><font color=red>Խ�����</font></div>')"
			Exit Sub
		End If
End If

Response.Write JsForm(ID)
End Sub

Call CloseConn()%>