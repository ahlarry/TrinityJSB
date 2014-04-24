<!--#include file="include/conn.asp"-->
<!--#include file="include/md5.asp"-->
<%
Dim action
action=LCase(Request("action"))
Select Case action
	Case "hlogin"
		Call HLogin()
	Case "logout"
		Call Logout()
	Case "login"
		Call Login()
	Case "indexlogin"
		Call IndexLogin()
	Case Else
		Call Main()
End Select

Sub Main()
	CurPage="用户登录"
	Call FileInc(0, "js/login.js")
	xjweb.Header()
	Call TopTablez()

	Dim userName, userPwd, SaveTime, LoginTime
	userName=Request.Cookies(web_info(10))("userName")
	userPwd=Request.Cookies(web_info(10))("userPwd")
	SaveTime=Request.Cookies(web_info(10))("saveTime")
	If Not IsNumeric(SaveTime) Then SaveTime=0
%>
	<Table class=xtable cellpadding=4 cellspacing=0 width="<%=web_info(8)%>"  align="center">
		<tr><td class=ctd Height=300>
			<Table class=ktable cellpadding=4 cellspacing=0 width=300 align="center">
				<form name=frm_login action="?action=login" method=post onsubmit="return login_true();">
				<tr>
					<td class=ctd colspan=2 height=30><font style=font-weight:bold;font-size:20px;>用户登录</font></td>
				</tr>
				<tr>
					<td class=rtd>用户名称:</td>
					<td class=ltd><input type=text name=userName size=15  value="<%=userName%>" style="background-image:url(images/login_bg.gif);background-position:right;background-repeat:no-repeat;"></td>
				</tr>
				<tr>
					<td class=rtd>用户密码:</td>
					<td class=ltd>
						<input type="password" name=userPwd size=15 value="<%=userPwd%>">
					</td>
				</tr>
				<tr>
					<td class=rtd>保存密码:</td>
					<td class=ltd>
						<Select name="SaveTime">
							<Option value=0 <%If SaveTime=0 Then Response.Write("Selected")%>>不保存</Option>
							<Option value=1 <%If SaveTime=1 Then Response.Write("Selected")%>>保存一天</Option>
							<Option value=31 <%If SaveTime=31 Then Response.Write("Selected")%>>保存一个月</Option>
							<Option value=365 <%If SaveTime=365 Then Response.Write("Selected")%>>保存一年</Option>
						</Select>
					</td>
				</tr>
				<tr>
					<td class=ctd colspan=2><input type=submit value=" 登 录 "></td>
				</tr>
				<input type="hidden" name="preUrl" value=<%=Request.ServerVariables("HTTP_REFFER")%>>
				</form>
			</table>
		</td></tr>
	</table>
<%
	Call BottomTable()
	xjweb.Footer()
	closeObj()
End Sub

Sub HLogin()
	If Session("userName")<>"" Then
		Rw("document.write('<b>"&Session("userName")&"</b>, 您好!  <a href=login.asp?action=logout>退出</a>');")
	Else
		Rw("document.write('欢迎使用本系统, 请先<a href=login.asp>登录</a>');")
	End If
End Sub

Sub Logout()
	Session("userName")=""
	Session("userAble")=NULL
	Response.Redirect("index.asp")
End Sub

Sub Login()
	Dim userName, userPwd, userIP, saveTime
	userName=Request("userName")
	userPwd=Request("userPwd")
	saveTime=Request("saveTime")
	userIP=xjweb.userip(0)
	strSql="Select * from [ims_user] where user_name='"&userName&"'"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.Eof Or Rs.Bof Then
		Call closeObj()
		Call JsAlert("用户 " & userName & " 不存在!请核实!","")
	Else
'		If Instr("朱磊aAAaa",userName)>0 and userIP<>"192.168.3.7" and userIP<>"127.0.0.1" Then
'			Call closeObj()
'			Call JsAlert("非法使用管理员帐户，请先获得授权！","")
'		End If
		If Rs("user_pwd")<>md5(userPwd,16) Then
			Call closeObj()
			Call JsAlert("密码不正确!请验证,并注意大小写!","")
		Else
			Session("userName")=userName
			Session("userPwd")=md5(userPwd,16)
			Session("userNick")=Rs("user_nick")
			Session("userAble")=Rs("user_able")
			Session("userdepart")=Rs("user_depart")
			Session("userGroup")=Rs("user_group")
			Session("userFace")=Rs("user_face")

			Response.Cookies(web_info(10)).Expires=Date + CInt(saveTime)
			Response.Cookies(web_info(10))("userName")=userName
			Response.Cookies(web_info(10))("userAble")=Rs("user_able")
			Response.Cookies(web_info(10))("userPwd")=userPwd
			Response.Cookies(web_info(10))("saveTime")=saveTime

			infoTitle="登录成功!"
			infoPreUrl=Request("preUrl")
			infoContents="欢迎登录 " & web_info(0) & "<br><br>" & AutoRefresh(3)
			Call GotoPrompt()
'			Response.Redirect("index.asp")
		End If
	End If
End Sub

Function TopTablez()
	Call SiteStat()	 '在此统计访问系统用户,有版权信息的页面均进行统计
%>
	<Div id="loading"  style=z-index:10000;visibility:hidden;position:'absolute';left:100;top:200;height:40;width:300;background-color:"#EEEEEE"; onclick="document.all.loading.style.visibility='hidden';">
		<Table cellpaddin=2 cellspacing=0 height="100%" width="100%" align="center">
			<tr><td align=center>Loading.......  Please Wait!</td></tr>
		</table>
	</Div>
	<Script language="javascript">
		document.all.loading.style.visibility='visible';
		document.all.loading.style.left=(screen.width-300)/2;
	</Script>
	<Table class=xtable width="<%=web_info(8)%>" cellpadding=0 cellspacing=0 border=0>
		<Tr><Td height=3 class=td_frame></td></Tr>
		<Tr><Td class=ctd height=22>
			<Table cellpadding=2 cellspacing=0 width="100%" height="100%">
				<tr>
					<td align=left width=350>&nbsp;&nbsp;Today: <%=XjDate(now,2)%></td>
					<td align=Right  width=*><script language="javascript" src="login.asp?action=hlogin"></script>&nbsp;&nbsp;</td>
				</tr>
			</table>
		</td></Tr>
		<Tr><Td class=ctd height=60>
			<Table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
				<tr><td align=center width=*><img src="<%=web_info(2)&web_info(9)%>"></td></tr>
			</table>
		</td></Tr>
	</Table>
	<%Response.Write(XjLine(2,web_info(8),""))%>

	<Table class=xtable width="<%=web_info(8)%>" cellpadding=0 cellspacing=0 border=0 align="center">
		<Tr><Td  class=ctd height=25>
			<Table border=0 cellpadding=5 cellspacing=0 width="100%" height="100%">
				<tr>
					<Td align=left width=*>◎<%=web_info(0) & " → " & CurPage%></Td>
					<Td align=center width=200><script language="javascript" src="inform_chk.asp?action=item"></script></Td>
					<% If ChkAble(0) Then %>
					<Td align=Right width=100><script language="javascript" src="msg_chk.asp?action=chknew"></script></Td>
					<% End If %>
				</tr>
			</table>
		</td></Tr>
	</Table>
	<%If strPage<>"" Then%>
	<%Response.Write(XjLine(2,web_info(8),""))%>
	<%End If%>
	 <%Response.Write(XjLine(2,web_info(8),""))%>
<%
End Function

%>