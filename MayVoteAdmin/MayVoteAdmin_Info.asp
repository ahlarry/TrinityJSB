<%
Dim InfoTitle,InfoContent,ComeUrl,InfoSub
Action = Request.QueryString("Action")
InfoContent = Request.QueryString("InfoContent")
ComeUrl = Request.QueryString("ComeUrl")
If Action = "OtherErr" Then
InfoTitle = "��������"
InfoSub = "<a href='javascript:history.go(-1)'>&lt;&lt;������һҳ</a>"
ElseIf Action = "Yes" Then
InfoTitle = "�����ɹ�<meta http-equiv='refresh' content='3;URL="&ComeUrl&"'>"
ElseIf Action = "LoginErr" Then
InfoTitle = "��δ��½��<meta http-equiv='refresh' content='3;URL=Admin_Login.asp'>"
Else
InfoTitle = "��������"
InfoContent = "<br><li>�벻Ҫֱ�ӷ��ʴ��ļ�</li>"
InfoSub = "<a href='javascript:window.close()'>�رմ���</a>"
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��Ϣ��ʾ - MayVoteͶƱϵͳ</title>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body><br><br><br><br><br><br>
<table width="400" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td height="25" background="Images/title.gif"><b><font color="#FFFFFF">����Ϣ��ʾ</font></b></td>
  </tr>
  <tr>
    <td align="center" bgcolor="#FFF3E6"><br>
      <table width="80%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td><b><% = InfoTitle%></b><br><% = InfoContent%></td>
        </tr>
    </table><br>
    <% = InfoSub%><br><br></td>
  </tr>
</table>
</body>
</html>