<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%'管理员验证
Call CheckUnAdmin()%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>投票项目添加 - MayVote投票系统</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><strong><font color="#FFFFFF">添
                加 投 票 项 目</font></strong></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td height="25" colspan="5"><b>　添加新项目</b>(标题在50字节以内)</td>
              </tr><form name="AddVote" method="post" action="Admin_AddVote_Posting.asp?Action=AddVote">
              <tr>
                <td width="15%" align="right"><b>名称：</b><br></td>
                <td colspan="3"><input name="MayVote_Topic" type="text" id="MayVote_Topic" size="40" maxlength="50"></td>
                <td width="20%" rowspan="3" align="center"><input type="submit" name="Submit" value="提交">　</td>
              </tr>
              <tr>
                <td align="right"><b>结束时间：</b></td>
                <td colspan="3"><select name="MayVote_EndDateTime1" id="MayVote_EndDateTime1">
                  <option value=<%=year(now)%> selected><%=year(now)%></option>
                  <option value=<%=year(now)+1%>><%=year(now)+1%></option>
                </select>
                年
                <select name="MayVote_EndDateTime2" id="MayVote_EndDateTime2">
                  <option value="01" selected>01</option>
                  <option value="02">02</option>
                  <option value="03">03</option>
                  <option value="04">04</option>
                  <option value="05">05</option>
                  <option value="06">06</option>
                  <option value="07">07</option>
                  <option value="08">08</option>
                  <option value="09">09</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
                </select>
                月
                <select name="MayVote_EndDateTime3" id="MayVote_EndDateTime3">
                  <option value="01" selected>01</option>
                  <option value="02">02</option>
                  <option value="03">03</option>
                  <option value="04">04</option>
                  <option value="05">05</option>
                  <option value="06">06</option>
                  <option value="07">07</option>
                  <option value="08">08</option>
                  <option value="09">09</option>
                  <option value="10">10</option>
                  <option value="11">11</option>
                  <option value="12">12</option>
                  <option value="13">13</option>
                  <option value="14">14</option>
                  <option value="15">15</option>
                  <option value="16">16</option>
                  <option value="17">17</option>
                  <option value="18">18</option>
                  <option value="19">19</option>
                  <option value="20">20</option>
                  <option value="21">21</option>
                  <option value="22">22</option>
                  <option value="23">23</option>
                  <option value="24">24</option>
                  <option value="25">25</option>
                  <option value="26">26</option>
                  <option value="27">27</option>
                  <option value="28">28</option>
                  <option value="29">29</option>
                  <option value="30">30</option>
                  <option value="31">31</option>
                </select>
                日</td>
                </tr>
              <tr>
                <td align="right"><b>单/多选：</b></td>
                <td width="30%"><select name="MayVote_Check" size="1" id="MayVote_Check">
                  <option value="0" selected>单选</option>
                  <option value="1">多选</option>
                </select></td>
                <td width="20%" align="right"><b>投票间隔：</b></td>
                <td><input name="MayVote_Time" type="text" id="MayVote_Time" value="240" size="3" maxlength="4">
小时</td>
                </tr>
              </form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><strong><font color="#FFFFFF">添
              加 投 票 选 项</font></strong></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td colspan="7"><b>　添加新选项</b>(标题在20字节以内)</td>
              </tr><form name="AddOption" method="post" action="Admin_AddVote_Posting.asp?Action=AddOption">
              <tr>
                <td width="10%" align="right"><b>名称：</b></td>
                <td width="20%" align="center"><input name="MayVote_Option" type="text" id="MayVote_Option" size="28" maxlength="20"></td>
                <td width="10%" align="right"><b>默认票数：</b></td>
                <td width="10%" align="center"><input name="MayVote_Nums" type="text" id="MayVote_Nums" value="0" size="5" maxlength="5"></td>
                <td width="10%" align="right"><b>所属项目：</b></td>
                <td width="25%" align="center"><select name="VID" size="1" id="VID">
                  <option value="0" selected></option>
<%
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select ID,MayVote_Topic From MayVote_Votes Order By ID Desc"
rs.Open SQL,Conn,1,1
While Not rs.Eof
If Len(rs("MayVote_Topic")) >12 Then
Response.Write"<option value='"&rs("ID")&"'>"&Left(MayHTMLEncode(rs("MayVote_Topic")),12)&"...</option>"
Else
Response.Write"<option value='"&rs("ID")&"'>"&MayHTMLEncode(rs("MayVote_Topic"))&"</option>"
End If
rs.MoveNext
Wend
rs.Close
Set rs = Nothing%>
                </select>                </td>
                <td width="10%"><input name="tj" type="submit" id="tj" value="提交"></td>
              </tr></form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
</body>
</html>
<%
Response.Write"<br><div align='center' class='smalltxt'>"
Call MayVote_CopyRight()
Response.Write"</div>"
Call CloseConn()
%>