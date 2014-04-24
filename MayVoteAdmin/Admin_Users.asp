<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%'管理员验证
Call CheckUnAdmin()

Action = Request.QueryString("Action")
If Action = "EditPassWord" Then
	Call EditPassWord()
Else
	Call Main()
End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户管理 - MayVote投票系统</title>
<link href="Images/style.css" rel="stylesheet" type="text/css"></head>
<body>
<%
'编辑用户
Sub EditPassWord()
%>
<br><table width="80%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><font color="#FFFFFF"><b>修
              改 个 人 密 码</b>　（UID：<% = Session("UID")%>）
              
        </font></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <form name="editpassword" method="post" action="Admin_Users_Posting.asp?Action=EditPassWord">
			  <tr>
                <td width="40%" height="25" align="right"><b>用户名：</b></td>
                <td><input name="UserName" type="text" readonly id="UserName" value="<% = Session("UserName")%>" size="20" maxlength="20"></td>
              </tr>
			  <tr>
			    <td height="25" align="right"><b>旧密码：</b></td>
			    <td><input name="OldPassWord" type="password" id="OldPassWord" size="20" maxlength="20"></td>
			  </tr>
              <tr>
                <td height="25" align="right"><b>新密码：</b></td>
                <td><input name="NewPassWord" type="password" id="NewPassWord" size="20" maxlength="20"></td>
              </tr>
              <tr>
                <td height="25" align="right"><b>确认新密码：</b></td>
                <td><input name="NewPassWord2" type="password" id="NewPassWord2" size="20" maxlength="20"><input name="UID" type="hidden" id="UID" value="<% = Session("UID")%>"></td>
              </tr>
              <tr>
                <td height="30" colspan="2" align="center"> <input name="tj" type="submit" id="tj" value=" 更改 "></td>
              </tr></form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
<%
End Sub
'默认页
Sub Main()
'禁止非超级管理员访问
Call CheckUnAdmin1()
%>
<script language="JavaScript">
function CheckAll(form)
{
	for (var i=0;i<form.elements.length;i++)
	{
		var e = form.elements[i];
		e.checked == true ? e.checked = false : e.checked = true;
	}
}
</script>
<br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><b><font color="#FFFFFF">添
                加 管 理 员</font></b></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td>
            <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <form name="form" method="post" action="Admin_Users_Posting.asp?Action=AddUsers"><tr>
                <td width="15%" height="40" align="right"><strong>用户名：</strong><br>
                  (4-20)字节以内</td>
                <td width="20%"><input name="UserName" type="text" id="UserName" size="20" maxlength="20"></td>
                <td width="10%" align="right"><strong>密&nbsp; 码：</strong><br>
                (4-16字节)</td>
                <td width="20%"><input name="PassWord" type="password" id="PassWord" size="20" maxlength="16"></td>
                <td width="12%"><select name="System" size="1">
                  <option value="0">普通管理员</option>
                  <option value="1">超级管理员</option>
                </select>
                </td>
                <td><input name="Submit" type="submit" id="Submit" value="提交"></td>
              </tr></form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table><br><form name="form2" method="post" action="Admin_Users_Posting.asp?Action=DelUsersAll">
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><strong><font color="#FFFFFF">管
              理 员 管 理</font></strong></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td width="5%" height="23" align="center" background="Images/admin_right_title02.gif">选定</td>
                <td width="10%" align="center" background="Images/admin_right_title02.gif">UID</td>
                <td width="35%" align="center" background="Images/admin_right_title02.gif">用户名</td>
                <td width="20%" align="center" background="Images/admin_right_title02.gif">权限</td>
                <td width="10%" align="center" background="Images/admin_right_title02.gif">状态</td>
                <td width="20%" align="center" background="Images/admin_right_title02.gif">常规操作</td>
              </tr><%
Set rs = Server.Createobject("adodb.Recordset")
SQL="Select UID,UserName,System,IsLock From May_Users"
rs.Open SQL,Conn,1,1
While Not rs.Eof
%><tr onMouseOver="this.bgColor='#FFDEAD'" onMouseOut="this.bgColor='#FFF3E6'">
                <td height="30" align="center"><%If rs("UID") <> Session("UID") Then Response.Write"<input type='checkbox' name='UID' value='"&rs("UID")&"'>"%></td>
                <td align="center"><% = rs("UID")%></td>
                <td align="center"><%If Session("UserName") = rs("UserName") Then
Response.Write "<font color='red'>"&MayHTMLEncode(rs("UserName"))&"</font>"
Else
Response.Write MayHTMLEncode(rs("UserName"))
End If%></td>
                <td align="center"><%If rs("System") = May_True Then
Response.Write"超级管理员"
Else
Response.Write"普通管理员"
End If%></td>
                <td align="center"><%If rs("IsLock") = May_False Then
Response.Write"<font color='green' title='正常'><b>√</b></font>"
Else
Response.Write"<font color='red' title='锁定'><b>×</b></font>"
End If%></td>
                <td align="center"><%If rs("UID") <> Session("UID") Then
If rs("IsLock") = May_False Then
Response.Write"<a href='Admin_Users_Posting.asp?Action=IsLock&UID="&rs("UID")&"' title='锁定'>锁定</a> | "
Else
Response.Write"<a href='Admin_Users_Posting.asp?Action=IsUnLock&UID="&rs("UID")&"' title='解锁'>解锁</a> | "
End If
End If
Response.Write"编辑"
If rs("UID") <> Session("UID") Then
Response.Write" | <a href='Admin_Users_Posting.asp?Action=DelUsers&UID="&rs("UID")&"' title='删除'>删除</a>"
End If%>
</td>
 </tr>
<%rs.MoveNext
Wend
rs.Close
Set rs = Nothing%>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table><div align="center"><br>
<input type='button' name='chkall' value='全选' onclick='CheckAll(this.form)'>
&nbsp; 
<input type=submit value='提交'>
&nbsp;&nbsp;&nbsp; 
<input name=action type=radio value='del'>
删除</div>
</form>
<%End Sub
Response.Write"<br><div align='center' class='smalltxt'>"
Call MayVote_CopyRight()
Response.Write"</div>"
Call CloseConn()
%></body>
</html>