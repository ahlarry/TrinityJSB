<!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MayVote_Function.asp"-->
<%'��Դ��֤
Call CheckUrl()
'����Ա��֤
Call CheckUnAdmin()
'��ֹ�ǳ�������Ա����
Call CheckUnAdmin1()

Action = Request.QueryString("Action")
If Action = "Updating" Then
	Call Updating()
Else
	Call Main()
End If
'���º�������
Sub Updating()
Dim MayVote_Name,MayVote_Url,MayVote_Setting,MayVote_Copy
MayVote_Name = Trim(Request.Form("MayVote_Name"))
MayVote_Url = Trim(Request.Form("MayVote_Url"))
MayVote_Setting = Trim(Request.Form("MayVote_Setting"))
MayVote_Copy = Trim(Request.Form("MayVote_Copy"))
If MayVote_Name = "" Or Len(MayVote_Name) > 30 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��վ����Ϊ�գ����䳤�ȴ���30�ֽڡ�</li>&Action=OtherErr"
If MayVote_Url = "" Or Len(MayVote_Url) > 50 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��վ��ַΪ�գ����䳤�ȴ���50�ֽڡ�</li>&Action=OtherErr"
If Len(MayVote_Copy) >255 Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û��Զ����Ȩ��ϢΪ�գ����䳤�ȴ���255�ֽڡ�</li>&Action=OtherErr"
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
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�������ø��³ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub
'��ҳ�� 
Sub Main()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�������� - MayVote��̨����</title>
<link href="Images/style.css" rel="stylesheet" type="text/css">
</head>
<body>
<br>
<table width="99%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#666666">
  <tr>
    <td><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td height="25" align="center" background="Images/title.gif"><div class="smalltxt"><b><font color="#FFFFFF">MayVote
                ϵ ͳ �� �� �� �� </font></b></div></td>
      </tr>
    </table>
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFF3E6">
          <tr>
            <td><table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#FFFFFF">
			<form name="form" method="post" action="Admin_Setting.asp?Action=Updating" target="_main">
              <tr>
                <td width="50%" height="28"><strong>��վ���ƣ�</strong><br>
                  վ���ƣ�����ʾ��ҳ��ײ�����ϵ��ʽ��</td>
                <td><input name="MayVote_Name" type="text" id="MayVote_Name" value="<% = Application("MayVote_Name")%>" size="30" maxlength="30"></td>
              </tr>
              <tr>
                <td width="50%" height="28"><strong>ͶƱϵͳ��ַ��</strong><br>
                  ����JS����</td>
                <td><input name="MayVote_Url" type="text" id="MayVote_Url" value="<% = Application("MayVote_Url")%>" size="30" maxlength="50"></td>
              </tr>
              <tr>
                <td width="50%" height="28"><strong>JS ��·���ƣ�</strong><br>
                 Ϊ�˱���������վ�Ƿ�������̳���ݣ��������ķ������������������������������̳ JS ����·�����б�ֻ�����б��е���������վ������ͨ�� JS ��������̳����Ϣ��<font color="#FF0000">ÿ������һ��</font>����֧��ͨ������������ http:// ���������������ݣ�����Ϊ��������·�����κ���վ���ɵ���)</td>
                <td><textarea name="MayVote_Setting" cols="30" rows="5" type="text" id="MayVote_Setting"><% = Application("MayVote_Setting")%></textarea></td>
              </tr>
              <tr>
                <td width="50%"><strong>�û���Ȩ��Ϣ��</strong><br>
                  ��ʾ��ͶƱ��ϸҳ�ĵײ�(֧��HTML�﷨)</td>
                <td><textarea name="MayVote_Copy" cols="30" rows="5" id="MayVote_Copy"><% = Application("MayVote_Copy")%>
                </textarea></td>
              </tr>
              <tr>
                <td height="40" colspan="2" align="center" valign="middle"><input type="submit" name="Submit" value="�ύ">��������
                  <input name="cz" type="reset" id="cz" value="����"></td>
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