<!--#include file="include/function.asp"-->
<!--#include file="include/md5.asp"-->
<%
if not isdebug then chkableinf(1)
select case request("action")
	case "admin_left"
		call admin_left()
	case "admin_login"
		call admin_login()
	case "admin_main"
		call admin_main()
	case "admin_head"
		call admin_head()
	case else
		call main()
end select

sub main()
	if not(chkable(1)) or session("admin")="" then
		call admin_login()
	else
%>
<html>
<head>
<title>��<%=xujian_ims.site_name%>��--ϵͳ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<script language="javascript">if(top.location!=self.location) top.location=self.location;</script>
<frameset id="frame" cols="180,*" frameborder="NO" border="0" framespacing="0" rows="*">
  <frame name="leftFrame" scrolling="AUTO" noresize src="admin_index.asp?action=admin_left" marginwidth="0" marginheight="0">
<%if not(chkable(1)) or session("admin")="" then%>
  <frame name="main" src="admin_index.asp?action=admin_login" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%else%>
  <frame name="main" src="admin_index.asp?action=admin_main" scrolling="AUTO" NORESIZE frameborder="0" marginwidth="10" marginheight="10" border="no">
<%end if%>
</frameset>
</frameset>
<noframes><body>�����������֧�ֿ��</body></noframes>
</html>
<%
	end if
end sub

sub admin_left()
%>

<title><%=xujian_ims.site_name%>--ϵͳ����������</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:normal 12px;
SCROLLBAR-FACE-COLOR: #799AE1; SCROLLBAR-HIGHLIGHT-COLOR: #799AE1;
SCROLLBAR-SHADOW-COLOR: #799AE1; SCROLLBAR-DARKSHADOW-COLOR: #799AE1;
SCROLLBAR-3DLIGHT-COLOR: #799AE1; SCROLLBAR-ARROW-COLOR: #FFFFFF;
SCROLLBAR-TRACK-COLOR: #AABFEC;
}
table  { border:0px; }
td  { font:normal 12px;}
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  {}
.menu_title span { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; cursor:hand; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
function hidesubmenu(sid)
{
eval("submenu" + sid + ".style.display=\"none\";");
}
//for(var i = 0;i<6; i ++)
	//hidesubmenu(i)
</SCRIPT>

<BODY leftmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<table width=100% cellpadding=0 cellspacing=0 border=0 align=left>
<tr><td valign=top>
	<!--������嶥�����-->
	<table cellpadding=0 cellspacing=0 width=158 align=center>
		<tr>
			<td height=42 valign=bottom><img src="images/admin/title.gif" width=158 height=38></td>
		</tr>
	</table>
	<table cellpadding=0 cellspacing=0 width=158 align=center>
		<tr>
			<td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background=images/admin/title_bg_quit.gif>
				<span><a href="admin_index.asp" target=_top><b>ϵͳ������ҳ</b></a> | <a href=admin_logout.asp target=_top><b>�˳�</b></a></span>
			</td>
		</tr>
	</table>
	&nbsp;

	<!--������Ϣ-->
	<%
	dim imenu	'����˵��ĸ���,ÿ��һ��������1
	imenu=0
	imenu=imenu+1
	%>
	<table cellpadding=0 cellspacing=0 width=158 align=center>
		<tr>
			<td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/admin_left_1.gif" id=menuTitle1 onclick="showsubmenu(<%=imenu%>)">
				<span>������Ϣ</span>
			</td>
		</tr>
		<tr>
			<td style="display='';" id='submenu<%=imenu%>'>
				<div class=sec_menu style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=150>
				<TBODY>
					<tr><td height=5></td></tr>
					<tr><td height=20>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_setting.asp target=main>��������</a>
						 | <a href=admin_bulletin.asp target=main>�������</a></td></tr>
				<TBODY>
				</table>
				</div>
				<div  style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=135>
					<tr><td height=20></td></tr>
				</table>
				</div>
			</td>
		</tr>
	</table>

	<!--�û�����-->
	<%imenu=imenu+1%>
	<table cellpadding=0 cellspacing=0 width=158 align=center>
		<tr>
			<td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/admin_left_3.gif" id=menuTitle1 onclick="showsubmenu(<%=imenu%>)">
				<span>�û�����</span>
			</td>
		</tr>
		<tr>
			<td style="display:none;" id='submenu<%=imenu%>'>
				<div class=sec_menu style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=150>
				<TBODY>
					<tr><td height=5></td></tr>
					<tr><td height=20>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_user.asp?action=adduser target=main>����û�</a>
						 | <a href=admin_user.asp?action=deluser target=main>ɾ���û�</a>
						 <br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_user.asp?action=setuser target=main>�����û�</a>
						<br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_admin.asp?action=chgpwd target=main>��������</a>
						 | <a href=admin_admin.asp?action=adminman target=main>����Ա����</a>
					</td></tr>
				<TBODY>
				</table>
				</div>
				<div  style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=135>
					<tr><td height=20></td></tr>
				</table>
				</div>
			</td>
		</tr>
	</table>

	<!--������Ϣ--���ݿ��ѯ-->
	<%imenu=imenu+1%>
	<table cellpadding=0 cellspacing=0 width=158 align=center>
		<tr>
			<td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/admin_left_3.gif" id=menuTitle1 onclick="showsubmenu(<%=imenu%>)">
				<span>��ѯ����</span>
			</td>
		</tr>
		<tr>
			<td style="display:none;" id='submenu<%=imenu%>'>
				<div class=sec_menu style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=150>
				<TBODY>
					<tr><td height=5></td></tr>
					<tr><td height=20>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=depart target=main>�û�����</a>
						 | <a href=admin_query.asp?action=sbcj target=main>�豸����</a><br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=mtjg target=main>ģͷ�ṹ</a>
						 | <a href=admin_query.asp?action=dwmc target=main>��λ����</a><br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=dxjg target=main>���ͽṹ</a>
						 | <a href=admin_query.asp?action=dmmc target=main>��������</a><br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=sxjg target=main>ˮ��ṹ</a>
						 | <a href=admin_query.asp?action=mjcl target=main>ģ�߲���</a><br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=fzbl target=main>��ֵ����</a>
						 | <a href=admin_query.asp?action=jcjxh target=main>�������ͺ�</a><br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=ckfz target=main>�ο���ֵ</a>
						 | <a href=admin_query.asp?action=rdogg target=main>�ȵ�ż���</a><br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=lxrwlx target=main>������������</a>
						<br>
						<img alt src="images/admin/bullet.gif" border="0" width="15" height="20">
						<a href=admin_query.asp?action=mtljcc target=main>ģͷ���ӳߴ�</a>
					</td></tr>
				<TBODY>
				</table>
				</div>
				<div  style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=135>
					<tr><td height=20></td></tr>
				</table>
				</div>
			</td>
		</tr>
	</table>

	<!--�ļ�����ʼ-->
	<%imenu=imenu+1%>
	<table cellpadding=0 cellspacing=0 width=158 align=center>
		<tr>
			<td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/admin_left_8.gif" id=menuTitle1 onclick="showsubmenu(<%=imenu%>)">
				<span>�ļ�����</span>
			</td>
		</tr>
		<tr>
			<td style="display:none;" id='submenu<%=imenu%>'>
				<div class=sec_menu style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=150>
				<TBODY>
					<tr><td height=5></td></tr>
					<tr><td height=20><img alt src="images/admin/bullet.gif" border="0" width="15" height="20"><a href=admin_upUserface.asp target=main>test</a></td></tr>
				<TBODY>
				</table>
				</div>
				<div  style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=135>
					<tr><td height=20></td></tr>
				</table>
				</div>
			</td>
		</tr>
	</table>

	<!--��Ȩ��Ϣ-->
	<table cellpadding=0 cellspacing=0 width=158 align=center>
		<tr>
			<td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin/admin_left_9.gif" id=menuTitle1>
				<span>��Ȩ��Ϣ</span>
			</td>
		</tr>
		<tr>
			<td>
				<div class=sec_menu style="width:158">
				<table cellpadding=0 cellspacing=0 align=center width=138>
					<tr><td height=20>
						<br>�������: <br>
							&nbsp;&nbsp;<b><%=xujian_ims.site_author%></b><BR>
						��ϵQQ: 119891935<BR>
						Email: ahxujian@126.com
						<br><br>
					</td></tr>
				</table>
				</div>
			</td>
		</tr>
	</table>
	&nbsp;
	<!--����������-->
<%
end sub

sub admin_login()
	rem ����Ա��¼���
	xujian_ims.web_adminhead()
	if not chkable(1) then response.write(prompt("������ϵͳ����Ա���<a href=login.asp>��¼</a>!�ٽ���ϵͳ����!"))

	if request("login")="chklogin" then
		call loginchk
	else
		call admin_login_main()
	end if
	xujian_ims.web_adminfoot()
end sub

sub admin_login_main()
	%>
	<table border="0" cellpadding="0" cellspacing="0" width="100%" height="60%">
		<tr>
			<td valign="middle" align="center">
				<table border="0" cellpadding="4" cellspacing="0" class="table_blue" width="280">
					<form action="" method="post" name="admin_login" onsubmit="return chkinf();">
					<tr>
						<td class="td_blue" colspan="2" height="35"><font style="font-size:14px; font-weight:bold;">����Ա��¼</font></td>
					</tr>
					<tr>
						<td class="td_rblue">����Ա</td>
						<td class="td_lblue"><input type="text" name="adminname" size="20"></td>
					</tr>
					<tr>
						<td class="td_rblue">��&nbsp;&nbsp;&nbsp;��</td>
						<td class="td_lblue"><input type="password" name="adminpwd" size="20"></td>
					</tr>
					<tr>
						<td class="td_cblue" colspan="2"><input type="submit" value=" �� ¼ "></td>
					</tr>
					<input type="hidden" name="login" value="chklogin">
					</form>
				</table>
			</td>
		</tr>
	</table>
	<script language="javascript">
	function chkinf()
	{
		if (document.admin_login.adminname.value=="")
			{alert("���������Ա���� ��"); document.admin_login.adminname.focus(); return false;}
		if (document.admin_login.adminpwd.value=="")
			{alert("���������Ա���� ��"); document.admin_login.adminpwd.focus(); return false;}
	 }
	 </script>
<%
end sub

sub loginchk()
	dim adminname, adminpwd
	adminname=trim(request("adminname"))
	adminpwd=md5(trim(request("adminpwd")),16)
	if adminname="" or adminpwd="" then
		response.write("<script language=""javascript"">alert('�û��������벻��Ϊ��!���ʵ������!');history.go(-1);</script>")
		exit sub
	end if
	sql="select * from admin where admin_name='"&adminname&"' and admin_pwd='"&adminpwd&"'"
	set rs=xujian_ims.exec(sql, 1)
	if rs.eof or rs.bof then
		rs.close
		response.write("<script language=""javascript"">alert('�û��������벻��ȷ,���ʵ������!');history.go(-1);</script>")
	else
		session("admin")=rs("admin_name")
		sql = "update admin set admin_lastlogin='"&now()&"',admin_lastloginip='"&xujian_ims.userip(0)&"' where admin_name='"&adminname&"'"
		call xujian_ims.exec(sql, 0)
		rs.close
		response.redirect "admin_index.asp"
	end if
end sub

sub admin_main()
	xujian_ims.web_adminhead()
	Dim theInstalledObjects(20)
    theInstalledObjects(0) = "MSWC.AdRotator"
    theInstalledObjects(1) = "MSWC.BrowserType"
    theInstalledObjects(2) = "MSWC.NextLink"
    theInstalledObjects(3) = "MSWC.Tools"
    theInstalledObjects(4) = "MSWC.Status"
    theInstalledObjects(5) = "MSWC.Counters"
    theInstalledObjects(6) = "IISSample.ContentRotator"
    theInstalledObjects(7) = "IISSample.PageCounter"
    theInstalledObjects(8) = "MSWC.PermissionChecker"
    theInstalledObjects(9) = "Scripting.FileSystemObject"
    theInstalledObjects(10) = "adodb.connection"

    theInstalledObjects(11) = "SoftArtisans.FileUp"
    theInstalledObjects(12) = "SoftArtisans.FileManager"
    theInstalledObjects(13) = "JMail.SMTPMail"	'Jamil 4.2
    theInstalledObjects(14) = "CDONTS.NewMail"
    theInstalledObjects(15) = "Persits.MailSender"
    theInstalledObjects(16) = "LyfUpload.UploadFile"
    theInstalledObjects(17) = "Persits.Upload.1"
	theInstalledObjects(18) = "JMail.Message"		'Jamil 4.3
	theInstalledObjects(19) = "Persits.Upload"
	theInstalledObjects(20) = "SoftArtisans.FileUp"
	%>
	<br>
	<table cellpadding="3" cellspacing="0" border="0" class="table_blue" align=center width="98%">
		<tr><th class="td_blue" colspan=2 height=25>��<%=xujian_ims.site_name%>�� ��Ϣͳ��</th></tr>
		<tr>
			<td width="50%"  class="td_lblue">
				���������ͣ�<%=Request.ServerVariables("OS")%>(IP:<%=Request.ServerVariables("LOCAL_ADDR")%>)
			</td>
			<td width="50%" class="td_lblue">
				�ű��������棺<%=ScriptEngine & "/"& ScriptEngineMajorVersion &"."&ScriptEngineMinorVersion&"."& ScriptEngineBuildVersion %>
			</td>
		</tr>
		<tr>
			<td width="50%" class="td_lblue" height=23>
				վ������·����<%=request.ServerVariables("APPL_PHYSICAL_PATH")%>
			</td>
			<td width="50%" class="td_lblue">
				���ݿ��ַ��<%=dbpath%>
			</td>
		</tr>
		<tr>
			<td width="50%" class="td_lblue" height=23>
				FSO�ı���д
				<%If Not IsObjInstalled(theInstalledObjects(9)) Then%>
					<font color=red><b>��</b></font>
				<%else%>
					<b>��</b>
				<%end if%>
			</td>
			<td width="50%" class="td_lblue">
				���ݿ�ʹ�ã�
				<%If Not IsObjInstalled(theInstalledObjects(10)) Then%>
					<font color=red><b>��</b></font>
				<%else%>
					<b>��</b>
				<%end if%>
			</td>
		</tr>
		<tr>
			<td width="50%" class="td_lblue" height=23>
				<%If IsObjInstalled(theInstalledObjects(18)) Then%>Jmail4.3�������֧�֣�<%else%>Jmail4.2���֧�֣�<%end if%>
				<%If IsObjInstalled(theInstalledObjects(18)) or IsObjInstalled(theInstalledObjects(13)) Then%>
					<b>��</b>
				<%else%>
					<font color=red><b>��</b></font>
				<%end if%>
			</td>
			<td width="50%" class="td_lblue">CDONTS�������֧�֣�<%If Not IsObjInstalled(theInstalledObjects(14)) Then%><font color=red><b>��</b></font><%else%><b>��</b><%end if%></td>
		</tr>
</table>

	<%
	xujian_ims.web_adminfoot()
end sub

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function
%>