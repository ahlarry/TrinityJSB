<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%'����Ա��֤
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
<title>�û����� - MayVoteͶƱϵͳ</title>
<link href="Images/style.css" rel="stylesheet" type="text/css"></head>
<body>
<%
'�༭�û�
Sub EditPassWord()
%>
<br><table width="80%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><font color="#FFFFFF"><b>��
              �� �� �� �� ��</b>����UID��<% = Session("UID")%>��
              
        </font></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <form name="editpassword" method="post" action="Admin_Users_Posting.asp?Action=EditPassWord">
			  <tr>
                <td width="40%" height="25" align="right"><b>�û�����</b></td>
                <td><input name="UserName" type="text" readonly id="UserName" value="<% = Session("UserName")%>" size="20" maxlength="20"></td>
              </tr>
			  <tr>
			    <td height="25" align="right"><b>�����룺</b></td>
			    <td><input name="OldPassWord" type="password" id="OldPassWord" size="20" maxlength="20"></td>
			  </tr>
              <tr>
                <td height="25" align="right"><b>�����룺</b></td>
                <td><input name="NewPassWord" type="password" id="NewPassWord" size="20" maxlength="20"></td>
              </tr>
              <tr>
                <td height="25" align="right"><b>ȷ�������룺</b></td>
                <td><input name="NewPassWord2" type="password" id="NewPassWord2" size="20" maxlength="20"><input name="UID" type="hidden" id="UID" value="<% = Session("UID")%>"></td>
              </tr>
              <tr>
                <td height="30" colspan="2" align="center"> <input name="tj" type="submit" id="tj" value=" ���� "></td>
              </tr></form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table>
<%
End Sub
'Ĭ��ҳ
Sub Main()
'��ֹ�ǳ�������Ա����
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
        <td height="25" align="center" background="Images/title.gif"><b><font color="#FFFFFF">��
                �� �� �� Ա</font></b></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td>
            <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <form name="form" method="post" action="Admin_Users_Posting.asp?Action=AddUsers"><tr>
                <td width="15%" height="40" align="right"><strong>�û�����</strong><br>
                  (4-20)�ֽ�����</td>
                <td width="20%"><input name="UserName" type="text" id="UserName" size="20" maxlength="20"></td>
                <td width="10%" align="right"><strong>��&nbsp; �룺</strong><br>
                (4-16�ֽ�)</td>
                <td width="20%"><input name="PassWord" type="password" id="PassWord" size="20" maxlength="16"></td>
                <td width="12%"><select name="System" size="1">
                  <option value="0">��ͨ����Ա</option>
                  <option value="1">��������Ա</option>
                </select>
                </td>
                <td><input name="Submit" type="submit" id="Submit" value="�ύ"></td>
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
        <td height="25" align="center" background="Images/title.gif"><strong><font color="#FFFFFF">��
              �� Ա �� ��</font></strong></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
              <tr>
                <td width="5%" height="23" align="center" background="Images/admin_right_title02.gif">ѡ��</td>
                <td width="10%" align="center" background="Images/admin_right_title02.gif">UID</td>
                <td width="35%" align="center" background="Images/admin_right_title02.gif">�û���</td>
                <td width="20%" align="center" background="Images/admin_right_title02.gif">Ȩ��</td>
                <td width="10%" align="center" background="Images/admin_right_title02.gif">״̬</td>
                <td width="20%" align="center" background="Images/admin_right_title02.gif">�������</td>
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
Response.Write"��������Ա"
Else
Response.Write"��ͨ����Ա"
End If%></td>
                <td align="center"><%If rs("IsLock") = May_False Then
Response.Write"<font color='green' title='����'><b>��</b></font>"
Else
Response.Write"<font color='red' title='����'><b>��</b></font>"
End If%></td>
                <td align="center"><%If rs("UID") <> Session("UID") Then
If rs("IsLock") = May_False Then
Response.Write"<a href='Admin_Users_Posting.asp?Action=IsLock&UID="&rs("UID")&"' title='����'>����</a> | "
Else
Response.Write"<a href='Admin_Users_Posting.asp?Action=IsUnLock&UID="&rs("UID")&"' title='����'>����</a> | "
End If
End If
Response.Write"�༭"
If rs("UID") <> Session("UID") Then
Response.Write" | <a href='Admin_Users_Posting.asp?Action=DelUsers&UID="&rs("UID")&"' title='ɾ��'>ɾ��</a>"
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
<input type='button' name='chkall' value='ȫѡ' onclick='CheckAll(this.form)'>
&nbsp; 
<input type=submit value='�ύ'>
&nbsp;&nbsp;&nbsp; 
<input name=action type=radio value='del'>
ɾ��</div>
</form>
<%End Sub
Response.Write"<br><div align='center' class='smalltxt'>"
Call MayVote_CopyRight()
Response.Write"</div>"
Call CloseConn()
%></body>
</html>