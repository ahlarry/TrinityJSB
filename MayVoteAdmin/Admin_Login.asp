<%
Response.Expires = 0 
Response.Expiresabsolute = Now() - 1 
Response.AddHeader "pragma","no-cache" 
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-cache"
%><!-- #include file="Const.asp" -->
<!-- #include file="../MayVote_Conn.asp" -->
<!-- #include file="Include/MD5.asp"-->
<!-- #include file="Include/MayVote_Function.asp"-->
<%
Action = Trim(Request("Action"))
If Action = "Login" then
	Call ChkLogin()
ElseIf Action = "Logout" then
	Call Logout()
Else
	Call Main()
End If
'��ҳ��
Sub Main()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>MayVote - �����½</title>
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("�������û�����");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("���������룡");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("������������֤�룡");
       document.Login.CheckCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("ϵͳ��ʾ��\n    ��ʹ�õ���Netscape����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("ϵͳ��ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  }
}
//-->
</script>
<style type="text/css">
<!--
table{ border-collapse: collapse;} 
A { TEXT-DECORATION: none; Color: #000000 }
A:hover { TEXT-DECORATION: underline;Color: #4455aa }
BODY { FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma; text-align: center; 
scrollbar-Base-Color: #1458DF;
SCROLLBAR-TRACK-COLOR: #EBEBEB;
SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;
SCROLLBAR-FACE-COLOR:#CCCCCC; 
SCROLLBAR-SHADOW-COLOR: #fff3e6;
SCROLLBAR-3DLIGHT-COLOR: yellow;
SCROLLBAR-DARKSHADOW-COLOR: red;
SCROLLBAR-ARROW-COLOR: #FFFFFF;
}
	font { line-height: normal; }
	TD { font-family: Tahoma; font-size: 12px; line-height:15px; }
	th { background-image:  url(Images/login_title.gIf); background-color: #4455aa; color: #D2691E; font-size: 12px; font-weight: bold; height:25;}
	th a { COLOR: #FFFFFF; TEXT-DECORATION: none; }
	th a:hover { COLOR: #FFFFFF; TEXT-DECORATION: underline; }


input,select,Textarea,option{ 
font-family: Tahoma,Verdana,"����"; font-size: 12px; line-height: 15px; COLOR: #000000;border: 1px #000000 solid}
-->
</style>
</head>
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="400" align="center" valign="middle"> 
      <table width="372" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr> 
          <th colspan="3" height=33><% = Application("MayVote_Name")%>
            �����½</th>
        </tr>
        <tr> 
          <td width="16"><img src="Images/login_left.gIf" width="16" height="187"></td>
          <td width="348"><table width="100%" height="175" border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td align="center" valign="top" bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0" background="Images/login_logobg.gIf">
                    <tr> 
                      <td height="15" colspan="2" align="right">&nbsp; </td>
                    </tr>
                    <tr> 
                      <td width=170 height="40"><a href="<% = Application("MayVote_Url")%>"></a></td>
                      <td width=><a href="<% = Application("MayVote_Url")%>"><b><% = Application("MayVote_Name")%></b></a><br>
                        �汾: MayVote <% Response.Write Application("MayVote_Ver")&" Acc��"%></td>
                    </tr>
                    <tr align="right" valign="middle"> 
                      <td height="20" colspan="2">&nbsp;</td>
                    </tr>
                  </table> 
                  <table bordercolor="#c8b09d" height="92" cellspacing="0" cellpadding="0" width="314" bgcolor="#fff3e6" border="1">
                    <tbody>
                      <tr> 
                        <td valign=center align=middle>
						<form action="Admin_Login.asp?Action=Login" method="post" name="Login" onSubmit="return CheckForm();">
                            <table cellspacing=0 cellpadding=0 width=314 border=0>
                              <tbody>
                                <tr> 
                                  <td width=244 height="20" align=left valign=center>&nbsp;&nbsp;�û�����
                                    <input name="UserName" type="text" id="UserName" size="20" maxlength="20"></td>
                                  <td width=70 rowspan="3" align=right valign=bottom>
								  <input name="Submit" type="image" style="width:60px; HEIGHT: 60px;border=0;" src="Images/login_logo.gIf" width="60" height="60"> 
                                  </td>
                                </tr>
                                <tr> 
                                  <td height="20" align=left valign=center>&nbsp;&nbsp;��&nbsp;&nbsp; 
                                    �룺
                                    <input name="PassWord" type="password" id="PassWord" size="20" maxlength="16"></td>
                                </tr>
                                <!--<tr> 
                                  <td height="20" align=left valign=center>&nbsp;&nbsp;��֤�룺
                                    <input name="GetCode" type="text" id="GetCode" size="20" maxlength="4">
                                    &nbsp;<img src="Include/May_GetCode.asp"><input name="Action" type="hidden" id="Action" value="Login"></td>
                                </tr>-->
                              </tbody>
                            </table>
						  </form>
						</td>
                      </tr>
                    </tbody>
                  </table>
                </td>
              </tr>
            </table>
            <img src="Images/login_bottom.gIf" width="348" height="12"></td>
          <td><img src="Images/Login_right.gIf" width="8" height="187"></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">
CheckBrowser();
SetFocus(); 
</script>
</body>
</html>
<%
End Sub
'��½
Sub ChkLogin()
Dim UserName,PassWord,CheckCode
UserName = ReplaceBadChar(MayHTMLEncode(Trim(Request.Form("UserName"))))
PassWord = Trim(Request.Form("PassWord"))
GetCode =ReplaceBadChar(Trim(Request.Form("GetCode")))
				
If UserName = "" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�����Ϊ�գ�</li>&Action=OtherErr"
If Password = "" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û�����Ϊ�գ�</li>&Action=OtherErr"
'If GetCode = "" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��֤��Ϊ�գ�</li>&Action=OtherErr"
'If Trim(Session("GetCode")) = "" Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��½��ʱ���뷵��ˢ��������д��</li>&Action=OtherErr"
'If GetCode<>CStr(Session("GetCode")) Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�������ȷ�����ϵͳ�����Ĳ�һ�£����������롣</li>&Action=OtherErr"
				
Password = md5(Password,16)
Set rs = Server.Createobject("adodb.RecordSet")
SQL = "Select UID,UserName,PassWord,System,IsLock From May_Users Where UserName='"&UserName&"'"
rs.Open SQL,conn,1,1
If rs.EOF And rs.BOF then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û������������</li>&Action=OtherErr"
Else
If PassWord <> rs("PassWord") Then
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û����ֻ��������</li>&Action=OtherErr"
	If rs("IsLock") = May_True Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û��Ѿ����������޷���½������ϵ����Ա</li>&Action=OtherErr"
Else
	If rs("IsLock") = May_True Then Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�û��Ѿ����������޷���½������ϵ����Ա</li>&Action=OtherErr"
Session("UID") = rs("UID")
Session("UserName") = rs("UserName")
Session("System") = rs("System")
End If
End If
rs.Close
Set rs = Nothing
Dim ComeUrl
ComeUrl = "Index.asp"
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>��½�ɹ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub

'�ǳ�
Sub Logout()
Session.Abandon()
Dim ComeUrl
ComeUrl = "Admin_Login.asp"
Response.Redirect "MayVoteAdmin_Info.asp?InfoContent=<br><li>�ɹ��ǳ���</li>&Action=Yes&ComeUrl="&ComeUrl&""
End Sub
Call CloseConn()
%>