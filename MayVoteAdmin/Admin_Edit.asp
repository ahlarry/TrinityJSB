<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
'��Դ��֤
Call CheckUrl()
'����Ա��֤
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
<title>�༭��Ŀ - MayVote��̨����</title>
<link href="Images/style.css" rel="stylesheet" type="text/css"></head>

<body>
<%Sub Main()%><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
<form name="VoteOrder" method="post" action="Admin_Edit_Posting.asp?Action=AllUpdate"><tr>
    <td align="center"><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><strong><font color="#FFFFFF">Ͷ
              Ʊ �� Ŀ �� ��</font></strong></td>
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
                <td>����<span class="smalltxt2">��</span>
                <a href="<% =Application("MayVote_Url")%>/MayVote.asp?Action=Show&ID=<% = rs("ID")%>" target="_blank">
                <%Response.Write"<b>"&MayHTMLEncode(rs("MayVote_Topic"))&"</b> (ͶƱʱ�䣺"&FormatDateTime(rs("MayVote_DateTime"),vbLongDate)&"&nbsp;-&nbsp;"&FormatDateTime(rs("MayVote_EndDateTime"),vbLongDate)&")"%></a> <span class="smalltxt2">- [</span><a href="Admin_Edit.asp?Action=EditVote&ID=<% = rs("ID")%>" target="main"><span class="smalltxt2">�༭</span></a><span class="smalltxt2">][</span><a href="Admin_Edit_Posting.asp?Action=DelVote&ID=<% = rs("ID")%>"><span class="smalltxt2">ɾ��</span></a><span class="smalltxt2">]</span></td>
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
                      <td><span class="smalltxt2">��</span> <% = MayHTMLEncode(rs1("MayVote_Option"))%> <span class="smalltxt2">- ��ʾ˳��</span>
                        <input name="MayVote_Order" type="text" value="<% = rs1("MayVote_Order")%>" size="1" maxlength="2">
                      - [<a href="Admin_Edit.asp?Action=EditOption&ID=<% = rs1("ID")%>">�༭</a>][<a href="Admin_Edit_Posting.asp?Action=DelOption&ID=<% = rs1("ID")%>">ɾ��</a>]
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
        <input name="gx" type="submit" id="tj" value=" ���� ">

     <input name="qc" type="submit" id="cz" value=" ��� "><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="5"></td>
          </tr>
        </table></td>
  </tr></form>
</table>
<%End Sub
'�༭ͶƱ��Ŀ��
Sub EditVote()
Dim ID
ID = Request.QueryString("ID")
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���ID������</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select ID,MayVote_Topic,MayVote_Check,MayVote_Time,MayVote_EndDateTime From MayVote_Votes Where ID="&ID
rs.Open SQL,Conn,1,1
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����ѡ���ͶƱ��Ŀ�����ڣ����Ѿ���ɾ����</li>&Action=OtherErr"
Else
%><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><b><font color="#FFFFFF">��
              �� Ͷ Ʊ �� Ŀ</font></b></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td height="25" colspan="4"><b>���༭ͶƱ��Ŀ</b>(������50�ֽ�����)</td>
              </tr>
              <form name="EditVote" method="post" action="Admin_Edit_Posting.asp?Action=EditVote">
                <tr>
                  <td align="right"><b>ID��</b><br></td>
                  <td colspan="3"><input name="ID" type="text" id="ID" value="<% = ID%>" size="4" maxlength="4" readonly></td>
                </tr>
                <tr>
                  <td align="right"><b>���ƣ�</b></td>
                  <td colspan="3"><input name="MayVote_Topic" type="text" id="MayVote_Topic" value="<% = rs("MayVote_Topic")%>" size="50" maxlength="50"></td>
                </tr>
                <tr>
                  <td align="right"><b>����ʱ�䣺</b></td>
                  <td colspan="3"><select name="MayVote_EndDateTime1" id="MayVote_EndDateTime1">
                    <option value=<%=year(now)-1%><%If Int(Year(rs("MayVote_EndDateTime"))) = year(now)-1 Then Response.Write"selected"%>><%=year(now)-1%></option>
                    <option value=<%=year(now)%><%If Int(Year(rs("MayVote_EndDateTime"))) = year(now) Then Response.Write"selected"%>><%=year(now)%></option>
                    <option value=<%=year(now)+1%><%If Int(Year(rs("MayVote_EndDateTime"))) = year(now)+1 Then Response.Write"selected"%>><%=year(now)+1%></option>
                  </select>
��
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
��
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
��</td>
                </tr>
                <tr>
                  <td width="20%" align="right"><b>��/��ѡ��</b></td>
                  <td width="30%"><select name="MayVote_Check" size="1" id="MayVote_Check">
                    <option value="0" <%If rs("MayVote_Check") = May_False Then Response.Write"selected"%>>��ѡ</option>
                    <option value="1" <%If rs("MayVote_Check") = May_True Then Response.Write"selected"%>>��ѡ</option>
                  </select></td>
                  <td width="20%" align="right"><b>ͶƱ�����</b></td>
                  <td width="30%"><input name="MayVote_Time" type="text" id="MayVote_Time" value="<% = rs("MayVote_Time")%>" size="3" maxlength="4">
Сʱ</td>
                </tr>
                <tr>
                  <td colspan="4" align="center"><input type="submit" name="Submit" value="����"></td>
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
'�༭ͶƱѡ��굥
Sub EditOption()
ID = Request.QueryString("ID")
If ID = "" Or isInteger(ID) = False Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�Ƿ���ID������</li>&Action=OtherErr"
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select VID,MayVote_Option,MayVote_Nums,MayVote_Order From MayVote_Options Where ID="&ID
rs.Open SQL,Conn,1,1
If rs.Eof And rs.Bof Then
	Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>����ѡ���ͶƱ��Ŀ�����ڣ����Ѿ���ɾ����</li>&Action=OtherErr"
Else
%><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><b><font color="#FFFFFF">��
              �� Ͷ Ʊ ѡ ��</font></b></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td colspan="4"><b>���༭ͶƱѡ��</b>(������20�ֽ�����)</td>
              </tr>
              <form name="EditOption" method="post" action="Admin_Edit_Posting.asp?Action=EditOption">
                <tr>
                  <td width="20%" align="right"><b>ID��</b></td>
                  <td colspan="3"><input name="ID" type="text" id="ID" value="<% = ID%>" size="4" maxlength="4" readonly></td>
                </tr>
                <tr>
                  <td align="right"><b>���ƣ�</b></td>
                  <td colspan="3"><input name="MayVote_Option" type="text" id="MayVote_Option" value="<% = rs("MayVote_Option")%>" size="28" maxlength="20"></td>
                </tr>
                <tr>
                  <td align="right"><b>Ĭ��Ʊ����</b></td>
                  <td width="30%"><input name="MayVote_Nums" type="text" id="MayVote_Nums" value="<% = rs("MayVote_Nums")%>" size="5" maxlength="5"></td>
                  <td width="20%" align="right"><b>��ʾ˳��</b></td>
                  <td><input name="MayVote_Order" type="text" id="MayVote_Order" value="<% = rs("MayVote_Order")%>" size="3" maxlength="4"></td>
                </tr>
                <tr>
                  <td align="right"><b>������Ŀ��</b></td>
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
                  <td colspan="4" align="center"><input name="tj" type="submit" id="tj" value="����"></td>
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
