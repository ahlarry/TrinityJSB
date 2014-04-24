<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%'来源验证
Call CheckUrl()
'管理员验证
Call CheckUnAdmin()
'禁止非超级管理员访问
Call CheckUnAdmin1()

Action = Request.QueryString("Action")
If Action = "Updating" Then
	Call Updating()
Else
	Call Main()
End If
'更新核心设置
Sub Updating()
Dim MayVote_Name,MayVote_Url,MayVote_Setting,MayVote_Copy
MayVote_Name = Trim(Request.Form("MayVote_Name"))
MayVote_Url = Trim(Request.Form("MayVote_Url"))
MayVote_Setting = Trim(Request.Form("MayVote_Setting"))
MayVote_Copy = Trim(Request.Form("MayVote_Copy"))
If MayVote_Name = "" Or Len(MayVote_Name) > 30 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>网站名称为空，或其长度大于30字节。</li>&Action=OtherErr"
If MayVote_Url = "" Or Len(MayVote_Url) > 50 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>网站地址为空，或其长度大于50字节。</li>&Action=OtherErr"
If Len(MayVote_Copy) >255 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>用户自定义版权信息为空，或其长度大于255字节。</li>&Action=OtherErr"
Set rs = Conn.Execute("Update MayVote_Config Set MayVote_Name = '"&MayVote_Name&"',MayVote_Url = '"&MayVote_Url&"',MayVote_Setting = '"&MayVote_Setting&"',MayVote_Copy = '"&MayVote_Copy&"' ")
Set rs = Nothing
Application.Lock
Application("MayVote_Name") = MayVote_Name
Application("MayVote_Url") = MayVote_Url
Application("MayVote_Setting") = MayVote_Setting
Application("MayVote_Copy") = MayVote_Copy
Application.UnLock
Dim ComeUrl
ComeUrl = Request.ServerVariables("HTTP_REFERER")
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>核心设置更新成功。</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub
'表单页面 
Sub Main()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>核心设置 - MayVote后台管理</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><div class="smalltxt"><b><font color="#FFFFFF">MayVote
                系 统 核 心 设 置 </font></b></div></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
			<form name="form" method="post" action="Admin_Setting.asp?Action=Updating" target="_main">
              <tr>
                <td width="50%" height="28"><strong>网站名称：</strong><br>
                  站名称，将显示在页面底部的联系方式处</td>
                <td><input name="MayVote_Name" type="text" id="MayVote_Name" value="<% = Application("MayVote_Name")%>" size="30" maxlength="30"></td>
              </tr>
              <tr>
                <td width="50%" height="28"><strong>投票系统地址：</strong><br>
                  用于JS调用</td>
                <td><input name="MayVote_Url" type="text" id="MayVote_Url" value="<% = Application("MayVote_Url")%>" size="30" maxlength="50"></td>
              </tr>
              <tr>
                <td width="50%" height="28"><strong>JS 来路限制：</strong><br>
                 为了避免其他网站非法调用论坛数据，加重您的服务器负担，您可以设置允许调用论坛 JS 的来路域名列表，只有在列表中的域名和网站，才能通过 JS 调用您论坛的信息。<font color="#FF0000">每个域名一行</font>，不支持通配符，请勿包含 http:// 或其他非域名内容，留空为不限制来路，即任何网站均可调用)</td>
                <td><textarea name="MayVote_Setting" cols="30" rows="5" type="text" id="MayVote_Setting"><% = Application("MayVote_Setting")%></textarea></td>
              </tr>
              <tr>
                <td width="50%"><strong>用户版权信息：</strong><br>
                  显示在投票详细页的底部(支持HTML语法)</td>
                <td><textarea name="MayVote_Copy" cols="30" rows="5" id="MayVote_Copy"><% = Application("MayVote_Copy")%>
                </textarea></td>
              </tr>
              <tr>
                <td height="40" colspan="2" align="center" valign="middle"><input type="submit" name="Submit" value="提交">　　　　
                  <input name="cz" type="reset" id="cz" value="重置"></td>
              </tr></form>
            </table></td>
          </tr>
      </table></td>
  </tr>
</table><br>
<div align="center" class="smalltxt"><%Call MayVote_CopyRight()
%>
</body>
</html>
<%End Sub
Call CloseConn()%>