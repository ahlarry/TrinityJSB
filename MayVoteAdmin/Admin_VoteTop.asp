<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
'来源验证
Call CheckUrl()
'管理员验证
Call CheckUnAdmin()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户管理 - MayVote投票系统</title>
<link href="Images/style.css" rel="stylesheet" type="text/css"></head>
<body>
<br><table width="80%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><font color="#FFFFFF"><b>投
              票 排 行 榜</b></font></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td>
            <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td width="10%" height="20" align="center" background="Images/admin_right_title02.gif" class="title">ID</td>
                <td width="60%" align="center" background="Images/admin_right_title02.gif" class="title">投票主题</td>
                <td width="10%" align="center" background="Images/admin_right_title02.gif" class="title">是否到期</td>
                <td align="center" background="Images/admin_right_title02.gif" class="title">得票数</td>
              </tr>
              <tr>
<%
Set rs = Server.Createobject("adodb.Recordset")
Sql="Select Sum(a.MayVote_Nums)  As MayVote_Nums, A.VID, B.MayVote_Topic, B.MayVote_EndDateTime From MayVote_Options As A Left Join MayVote_Votes As B on A.VID=B.ID Group By A.VID, B.MayVote_Topic, B.MayVote_EndDateTime"
rs.Open Sql,Conn,1,1
While Not rs.Eof
%>
                <td height="20" align="center"><% = rs("VID")%></td>
                <td height="20">&nbsp;<a href="<% =Application("MayVote_Url")%>/MayVote.asp?Action=Show&ID=<% = rs("VID")%>" target="_blank"><% = rs("MayVote_Topic")%></a></td>
                <td height="20" align="center">
<%If Now() < rs("MayVote_EndDateTime") Then
	Response.Write"<font color=green>否</font>"
Else
	Response.Write"<font color=red>是</font>"
End If
%></td>
                <td height="20" align="center"><font color="#FF9900"><% = rs("MayVote_Nums")%></font></td>
              </tr><%rs.MoveNext
Wend
rs.Close
Set rs = Nothing%>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
<br>
<div align="center" class="smalltxt"><%Call MayVote_CopyRight()
Call CloseConn()
%></div>
</body>
</html>