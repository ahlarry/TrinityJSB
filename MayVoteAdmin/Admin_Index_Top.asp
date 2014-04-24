<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>TOP</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body leftmargin="0" topmargin="0">
<table width="100%" border="0" cellpadding="0" cellspacing="0" background="Images/top_bg.gif">
 <tr> 
<td width="150" height="35"></td>
<td width="20"></td>
<td><div class="smalltxt"><b><font color="#FFFFFF">欢迎 <% = Session("UserName")%> 
进入MayVote设置面板　　&nbsp; </font></b><b><font color="#FFFFFF"><a href="Admin_help.asp" target="main"><font color="#FFFFFF">使用帮助</font></a></font></b></div></td>
<td width="300" align="right"><%Response.Write "<div class='smalltxt'><b><font color='#FFFFFF'>MayVote"&Application("MayVote_Ver")&"&nbsp;&nbsp;&nbsp;&nbsp;</font></b></div>"%></td>
</tr>
</table>
</body>
</html>
