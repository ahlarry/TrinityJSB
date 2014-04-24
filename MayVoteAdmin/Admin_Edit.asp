<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
'来源验证
Call CheckUrl()
'管理员验证
Call CheckUnAdmin()
Action = Request.QueryString("Action")
If Action = "EditVote" Then
	Call EditVote()
ElseIf Action = "EditOption" Then
	Call EditOption()
Else
	Call Main()
End If%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>编辑项目 - MayVote后台管理</title>
<link href="Images/style.css" rel="stylesheet" type="text/css"></head>

<body>
<%Sub Main()%><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
<form name="VoteOrder" method="post" action="Admin_Edit_Posting.asp?Action=AllUpdate"><tr>
    <td align="center"><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><strong><font color="#FFFFFF">投
              票 项 目 管 理</font></strong></td>
      </tr>
    </table><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5"></td>
          </tr>
        </table>
<%
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select * From MayVote_Votes"
rs.Open SQL,Conn,1,1
While Not rs.Eof
%>
        <table width="98%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="5" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td>　　<span class="smalltxt2">●</span>
                <a href="<% =Application("MayVote_Url")%>/MayVote.asp?Action=Show&ID=<% = rs("ID")%>" target="_blank">
                <%Response.Write"<b>"&MayHTMLEncode(rs("MayVote_Topic"))&"</b> (投票时间："&FormatDateTime(rs("MayVote_DateTime"),vbLongDate)&"&nbsp;-&nbsp;"&FormatDateTime(rs("MayVote_EndDateTime"),vbLongDate)&")"%></a> <span class="smalltxt2">- [</span><a href="Admin_Edit.asp?Action=EditVote&ID=<% = rs("ID")%>" target="main"><span class="smalltxt2">编辑</span></a><span class="smalltxt2">][</span><a href="Admin_Edit_Posting.asp?Action=DelVote&ID=<% = rs("ID")%>"><span class="smalltxt2">删除</span></a><span class="smalltxt2">]</span></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
Set rs1 = Server.Createobject("adodb.Recordset")
SQL="Select ID,MayVote_Option,MayVote_Order From MayVote_Options Where VID="&rs("ID")&" Order By MayVote_Order Asc"
rs1.Open SQL,Conn,1,1
While Not rs1.Eof
%>
                    <tr>
                      <td width="10%" height="22"></td>
                      <td><span class="smalltxt2">○</span> <% = MayHTMLEncode(rs1("MayVote_Option"))%> <span class="smalltxt2">- 显示顺序：</span>
                        <input name="MayVote_Order" type="text" value="<% = rs1("MayVote_Order")%>" size="1" maxlength="2">
                      - [<a href="Admin_Edit.asp?Action=EditOption&ID=<% = rs1("ID")%>">编辑</a>][<a href="Admin_Edit_Posting.asp?Action=DelOption&ID=<% = rs1("ID")%>">删除</a>]
                      <input name="OptionID" type="hidden" id="OptionID" value="<% = rs1("ID")%>"></td>
                    </tr><%rs1.MoveNext
Wend
rs1.Close
Set rs1 = Nothing%>
                  </table></td>
              </tr>
            </table></td>
          </tr>
      </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5"></td>
          </tr>
        </table><%rs.MoveNext
Wend
rs.Close
Set rs = Nothing%>
        <input name="gx" type="submit" id="tj" value=" 更新 ">

     <input name="qc" type="submit" id="cz" value=" 清除 "><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5"></td>
          </tr>
        </table></td>
  </tr></form>
</table>
<%End Sub
'编辑投票项目表单
Sub EditVote()
Dim ID
ID = Request.QueryString("ID")
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的ID参数。</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select ID,MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_EndDateTime From MayVote_Votes Where ID="&ID
rs.Open SQL,Conn,1,1
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>你所选择的投票项目不存在，或已经被删除。</li>&Action=OtherErr"
Else
%><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><b><font color="#FFFFFF">编
              辑 投 票 项 目</font></b></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td height="25" colspan="4"><b>　编辑投票项目</b>(标题在50字节以内)</td>
              </tr>
              <form name="EditVote" method="post" action="Admin_Edit_Posting.asp?Action=EditVote">
                <tr>
                  <td align="right"><b>ID：</b><br></td>
                  <td colspan="3"><input name="ID" type="text" id="ID" value="<% = ID%>" size="4" maxlength="4" readonly></td>
                </tr>
                <tr>
                  <td align="right"><b>名称：</b></td>
                  <td colspan="3"><input name="MayVote_Topic" type="text" id="MayVote_Topic" value="<% = rs("MayVote_Topic")%>" size="50" maxlength="50"></td>
                </tr>
                <tr>
                  <td align="right"><b>结束时间：</b></td>
                  <td colspan="3"><select name="MayVote_EndDateTime1" id="MayVote_EndDateTime1">
                    <option value=<%=year(now)-1%><%If Int(Year(rs("MayVote_EndDateTime"))) = year(now)-1 Then Response.Write"selected"%>><%=year(now)-1%></option>
                    <option value=<%=year(now)%><%If Int(Year(rs("MayVote_EndDateTime"))) = year(now) Then Response.Write"selected"%>><%=year(now)%></option>
                    <option value=<%=year(now)+1%><%If Int(Year(rs("MayVote_EndDateTime"))) = year(now)+1 Then Response.Write"selected"%>><%=year(now)+1%></option>
                  </select>
年
<select name="MayVote_EndDateTime2" id="MayVote_EndDateTime2">
  <option value="01"<%If Int(Month(rs("MayVote_EndDateTime"))) = 1 Then Response.Write"selected"%>>01</option>
  <option value="02"<%If Int(Month(rs("MayVote_EndDateTime"))) = 2 Then Response.Write"selected"%>>02</option>
  <option value="03"<%If Int(Month(rs("MayVote_EndDateTime"))) = 3 Then Response.Write"selected"%>>03</option>
  <option value="04"<%If Int(Month(rs("MayVote_EndDateTime"))) = 4 Then Response.Write"selected"%>>04</option>
  <option value="05"<%If Int(Month(rs("MayVote_EndDateTime"))) = 5 Then Response.Write"selected"%>>05</option>
  <option value="06"<%If Int(Month(rs("MayVote_EndDateTime"))) = 6 Then Response.Write"selected"%>>06</option>
  <option value="07"<%If Int(Month(rs("MayVote_EndDateTime"))) = 7 Then Response.Write"selected"%>>07</option>
  <option value="08"<%If Int(Month(rs("MayVote_EndDateTime"))) = 8 Then Response.Write"selected"%>>08</option>
  <option value="09"<%If Int(Month(rs("MayVote_EndDateTime"))) = 9 Then Response.Write"selected"%>>09</option>
  <option value="10"<%If Int(Month(rs("MayVote_EndDateTime"))) = 10 Then Response.Write"selected"%>>10</option>
  <option value="11"<%If Int(Month(rs("MayVote_EndDateTime"))) = 11 Then Response.Write"selected"%>>11</option>
  <option value="12"<%If Int(Month(rs("MayVote_EndDateTime"))) = 12 Then Response.Write"selected"%>>12</option>
</select>
月
<select name="MayVote_EndDateTime3" id="MayVote_EndDateTime3">
  <option value="01"<%If Int(Day(rs("MayVote_EndDateTime"))) = 1 Then Response.Write"selected"%>>01</option>
  <option value="02"<%If Int(Day(rs("MayVote_EndDateTime"))) = 2 Then Response.Write"selected"%>>02</option>
  <option value="03"<%If Int(Day(rs("MayVote_EndDateTime"))) = 3 Then Response.Write"selected"%>>03</option>
  <option value="04"<%If Int(Day(rs("MayVote_EndDateTime"))) = 4 Then Response.Write"selected"%>>04</option>
  <option value="05"<%If Int(Day(rs("MayVote_EndDateTime"))) = 5 Then Response.Write"selected"%>>05</option>
  <option value="06"<%If Int(Day(rs("MayVote_EndDateTime"))) = 6 Then Response.Write"selected"%>>06</option>
  <option value="07"<%If Int(Day(rs("MayVote_EndDateTime"))) = 7 Then Response.Write"selected"%>>07</option>
  <option value="08"<%If Int(Day(rs("MayVote_EndDateTime"))) = 8 Then Response.Write"selected"%>>08</option>
  <option value="09"<%If Int(Day(rs("MayVote_EndDateTime"))) = 9 Then Response.Write"selected"%>>09</option>
  <option value="10"<%If Int(Day(rs("MayVote_EndDateTime"))) = 10 Then Response.Write"selected"%>>10</option>
  <option value="11"<%If Int(Day(rs("MayVote_EndDateTime"))) = 11 Then Response.Write"selected"%>>11</option>
  <option value="12"<%If Int(Day(rs("MayVote_EndDateTime"))) = 12 Then Response.Write"selected"%>>12</option>
  <option value="13"<%If Int(Day(rs("MayVote_EndDateTime"))) = 13 Then Response.Write"selected"%>>13</option>
  <option value="14"<%If Int(Day(rs("MayVote_EndDateTime"))) = 14 Then Response.Write"selected"%>>14</option>
  <option value="15"<%If Int(Day(rs("MayVote_EndDateTime"))) = 15 Then Response.Write"selected"%>>15</option>
  <option value="16"<%If Int(Day(rs("MayVote_EndDateTime"))) = 16 Then Response.Write"selected"%>>16</option>
  <option value="17"<%If Int(Day(rs("MayVote_EndDateTime"))) = 17 Then Response.Write"selected"%>>17</option>
  <option value="18"<%If Int(Day(rs("MayVote_EndDateTime"))) = 18 Then Response.Write"selected"%>>18</option>
  <option value="19"<%If Int(Day(rs("MayVote_EndDateTime"))) = 19 Then Response.Write"selected"%>>19</option>
  <option value="20"<%If Int(Day(rs("MayVote_EndDateTime"))) = 20 Then Response.Write"selected"%>>20</option>
  <option value="21"<%If Int(Day(rs("MayVote_EndDateTime"))) = 21 Then Response.Write"selected"%>>21</option>
  <option value="22"<%If Int(Day(rs("MayVote_EndDateTime"))) = 22 Then Response.Write"selected"%>>22</option>
  <option value="23"<%If Int(Day(rs("MayVote_EndDateTime"))) = 23 Then Response.Write"selected"%>>23</option>
  <option value="24"<%If Int(Day(rs("MayVote_EndDateTime"))) = 24 Then Response.Write"selected"%>>24</option>
  <option value="25"<%If Int(Day(rs("MayVote_EndDateTime"))) = 25 Then Response.Write"selected"%>>25</option>
  <option value="26"<%If Int(Day(rs("MayVote_EndDateTime"))) = 26 Then Response.Write"selected"%>>26</option>
  <option value="27"<%If Int(Day(rs("MayVote_EndDateTime"))) = 27 Then Response.Write"selected"%>>27</option>
  <option value="28"<%If Int(Day(rs("MayVote_EndDateTime"))) = 28 Then Response.Write"selected"%>>28</option>
  <option value="29"<%If Int(Day(rs("MayVote_EndDateTime"))) = 29 Then Response.Write"selected"%>>29</option>
  <option value="30"<%If Int(Day(rs("MayVote_EndDateTime"))) = 30 Then Response.Write"selected"%>>30</option>
  <option value="31"<%If Int(Day(rs("MayVote_EndDateTime"))) = 31 Then Response.Write"selected"%>>31</option>
</select>
日</td>
                </tr>
                <tr>
                  <td width="20%" align="right"><b>单/多选：</b></td>
                  <td width="30%"><select name="MayVote_Check" size="1" id="MayVote_Check">
                    <option value="0" <%If rs("MayVote_Check") = May_False Then Response.Write"selected"%>>单选</option>
                    <option value="1" <%If rs("MayVote_Check") = May_True Then Response.Write"selected"%>>多选</option>
                  </select></td>
                  <td width="20%" align="right"><b>投票间隔：</b></td>
                  <td width="30%"><input name="MayVote_Time" type="text" id="MayVote_Time" value="<% = rs("MayVote_Time")%>" size="3" maxlength="4">
小时</td>
                </tr>
                <tr>
                  <td colspan="4" align="center"><input type="submit" name="Submit" value="更新"></td>
                </tr>
              </form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
<%
End If
rs.Close
Set rs = Nothing
End Sub
'编辑投票选项标单
Sub EditOption()
ID = Request.QueryString("ID")
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>非法的ID参数。</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select VID,MayVote_Option,MayVote_Nums,MayVote_Order From MayVote_Options Where ID="&ID
rs.Open SQL,Conn,1,1
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>你所选择的投票项目不存在，或已经被删除。</li>&Action=OtherErr"
Else
%><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><b><font color="#FFFFFF">编
              辑 投 票 选 项</font></b></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td colspan="4"><b>　编辑投票选项</b>(标题在20字节以内)</td>
              </tr>
              <form name="EditOption" method="post" action="Admin_Edit_Posting.asp?Action=EditOption">
                <tr>
                  <td width="20%" align="right"><b>ID：</b></td>
                  <td colspan="3"><input name="ID" type="text" id="ID" value="<% = ID%>" size="4" maxlength="4" readonly></td>
                </tr>
                <tr>
                  <td align="right"><b>名称：</b></td>
                  <td colspan="3"><input name="MayVote_Option" type="text" id="MayVote_Option" value="<% = rs("MayVote_Option")%>" size="28" maxlength="20"></td>
                </tr>
                <tr>
                  <td align="right"><b>默认票数：</b></td>
                  <td width="30%"><input name="MayVote_Nums" type="text" id="MayVote_Nums" value="<% = rs("MayVote_Nums")%>" size="5" maxlength="5"></td>
                  <td width="20%" align="right"><b>显示顺序：</b></td>
                  <td><input name="MayVote_Order" type="text" id="MayVote_Order" value="<% = rs("MayVote_Order")%>" size="3" maxlength="4"></td>
                </tr>
                <tr>
                  <td align="right"><b>所属项目：</b></td>
                  <td colspan="3"><select name="VID" size="1" id="VID">
<%
Dim Selected
Set rs1 = Server.Createobject("adodb.Recordset")
SQL="Select ID,MayVote_Topic From MayVote_Votes Order By ID Desc"
rs1.Open SQL,Conn,1,1
While Not rs1.Eof
If rs("VID") = rs1("ID") Then
Selected = "selected"
Else
Selected = ""
End If
If Len(rs1("MayVote_Topic")) >15 Then
Response.Write"<option value='"&rs1("ID")&"' "&Selected&">"&Left(MayHTMLEncode(rs1("MayVote_Topic")),10)&"...</option>"
Else
Response.Write"<option value='"&rs1("ID")&"' "&Selected&">"&Left(MayHTMLEncode(rs1("MayVote_Topic")),15)&"</option>"
End If
rs1.MoveNext
Wend
rs1.Close
Set rs1 = Nothing%>
                  </select></td>
                </tr>
                <tr>
                  <td colspan="4" align="center"><input name="tj" type="submit" id="tj" value="更新"></td>
                </tr>
              </form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
<%
End If
rs.Close
Set rs = Nothing
End Sub
Response.Write"<br><div align='center' class='smalltxt'>"
Call MayVote_CopyRight()
Response.Write"</div>"
Call CloseConn()
%>
</body>
</html>
