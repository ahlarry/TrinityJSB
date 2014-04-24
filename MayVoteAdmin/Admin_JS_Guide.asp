<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
'来源验证
Call CheckUrl()
'管理员验证
Call CheckUnAdmin()

Dim VID,Content
VID = Request.Form("VID")
If VID = "" Then
Content = ""
Else
Content = "&lt;script language=&quot;JavaScript&quot; src=&quot;MayVote.asp?Action=JS&ID="&VID&"&quot;&gt;&lt;/script&gt;"
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>JS调用向导 - MayVote后台管理</title>
<link href="Images/style.css" rel="stylesheet" type="text/css"></head>

<body><br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><b><font color="#FFFFFF">JS
              调 用 向 导</font></b></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
            <form name="JS" method="post" action="Admin_JS_Guide.asp">
			<tr>
                <td width="100%" colspan="2" align="center"><br><b>插入以下代码在需要调用的地方即可</b><br><br><textarea rows="4" style="width: 80%; word-break: break-all" onMouseOver="this.focus()" onFocus="this.select()" name"Content"><% = Content%></textarea></td>
                </tr>
                <tr>
                  <td width="30%" height="25" align="right"><b>请选择您需要调用的调查项目：</b><br>(注意：这里只可以调用未到期的投票项目)</td>
                  <td width="75%"><select name="VID" size="5" id="VID">
<%
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select ID,MayVote_Topic From MayVote_Votes Where Now() < MayVote_EndDateTime Order By ID Desc"
rs.Open SQL,Conn,1,1
While Not rs.Eof
If Len(rs("MayVote_Topic")) >30 Then
Response.Write"<option value='"&rs("ID")&"'>"&Left(MayHTMLEncode(rs("MayVote_Topic")),30)&"...</option>"
Else
Response.Write"<option value='"&rs("ID")&"'>"&MayHTMLEncode(rs("MayVote_Topic"))&"</option>"
End If
rs.MoveNext
Wend
rs.Close
Set rs = Nothing%>
                  </select>                  </td>
                </tr>
                  <tr>
                    <td colspan="2" align="center"><input type="submit" name="Submit" value="提交"></td>
                  </tr>
                </form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
<%
Response.Write"<br><div align='center' class='smalltxt'>"
Call MayVote_CopyRight()
Response.Write"</div>"
Call CloseConn()
%>
</body>
</html>