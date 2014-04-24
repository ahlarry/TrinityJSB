<!-- #include file="Const.asp" -->
<%
'管理员验证
Call CheckUnAdmin()
%><html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>MayVote - 后台管理</title>
</head>

<frameset rows="35,*" cols="*" frameborder="NO" border="0" framespacing="0">
  <frame src="Admin_Index_Top.asp" name="top" scrolling="NO" noresize >
<frameset rows="*" cols="181,*" framespacing="0" frameborder="0" border="false" id="frame" scrolling="yes">
  <frame name="left" scrolling="auto" marginwidth="0" marginheight="0" src="Admin_Index_Left.asp">
    <frame src="Admin_Index_Main.asp" name="main">
  </frameset>
</frameset>
<noframes><body>

</body></noframes>
</html>