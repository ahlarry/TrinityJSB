<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<%'����Ա��֤
Call CheckUnAdmin()
%>
<html>
<head>
<title>MayVote--��̨������ҳ</title>
<meta http-equiv="Content-Type" content="text/html; charSet =gb2312">
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#F6F6F6" leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td> <table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="25" align="center" background="Images/title.gif"><div class="smalltxt"><b><font color="#FFFFFF">MayVote
                �� Ϣ ͳ ��</font></b></div></td>
        </tr>
      </table>
      <table width="100%" border="1" align=center cellpadding="3" cellspacing="0" bordercolor="#FFFFFF" class="border">
        <tr bgcolor="#FFF3E6"> 
          <td height=23 colspan="2" background="Images/admin_right_title02.gif" class="title">&nbsp;&nbsp;������Ŀ������<b><font color="#FF0000">
            <%
Set rs = Conn.Execute("Select Count(ID) From MayVote_Votes")
Response.Write rs(0)
Set rs = Nothing%>
          </font></b> ��ѡ��������<b><font color="#FF0000">
          <%Set rs = Conn.Execute("Select Count(ID) From MayVote_Options")
Response.Write rs(0)
Set rs = Nothing%>
          </font></b> ��ͶƱ������<b><font color="red">
          <%Set rs = Conn.Execute("Select Sum(MayVote_Nums) As Nums From MayVote_Options")
Response.Write rs("Nums")
Set rs = Nothing%></font></b>
        <tr bgcolor="#FFCC66"> 
          <td height=25 colspan="2" bgcolor="#FFCC00" class="title">&nbsp;&nbsp;����ʽ��Ȩ�� 
            <font color="#FF0000"><% = Application("MayVote_Name")%>
            </font> ʹ�ã���ǰʹ�ð汾Ϊ <font color="#FF0000">MayVote
            <% = Application("MayVote_Ver")%></font>
        <tr bgcolor="#FFF3E6"> 
          <td height=23 bgcolor="#FFF3E6" class="title">&nbsp;&nbsp;���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>) 
          <td width="50%" bgcolor="#FFF3E6" class="title">�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %> 
        <tr bgcolor="#FFF3E6"> 
          <td width="50%" height=23 class="tdbg">&nbsp;&nbsp;վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
          <td width="50%" class="tdbg">&nbsp;</td>
        </tr>
      </table> </td>
  </tr>
</table>
<div align="center" class="smalltxt"><%Call MayVote_CopyRight()
Call CloseConn()
%></div>
</body>
</html>
