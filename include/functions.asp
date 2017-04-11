<%
Function TopTable()
	Call SiteStat()	 '在此统计访问系统用户,有版权信息的页面均进行统计
%>
	<Div id="loading"  style=z-index:10000;visibility:hidden;position:'absolute';left:100;top:200;height:40;width:300;background-color:"#EEEEEE"; onclick="document.all.loading.style.visibility='hidden';">
		<Table cellpaddin=2 cellspacing=0 height="100%" width="100%">
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
		<Tr><Td  class=ctd height=22>
			<Table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
				<tr><td align=center width=*><%=mainmenu%></td></tr>
			</table>
		</td></Tr>
	</Table>

	<%Response.Write(XjLine(2,web_info(8),""))%>

	<Table class=xtable width="<%=web_info(8)%>" cellpadding=0 cellspacing=0 border=0>
		<Tr><Td  class=ctd height=25>
			<Table border=0 cellpadding=5 cellspacing=0 width="100%" height="100%">
				<tr>
					<Td align=left width=*>◎<%=web_info(0) & " → " & CurPage%></Td>
					<% iF ChkAble(0) Then %>
					<Td align=center width=200><script language="javascript" src="inform_chk.asp?action=item"></script></Td>
					<Td align=Right><a href="..\dbm" target="_blank"><font style="font-size:18px;" face="华文行楷" color="#ff8080"><strong><span style="background-color: #ffff80">模具信息库测试</span></strong></font></a></Td>
					<Td align=Right width=*><script language="javascript" src="msg_chk.asp?action=chknew"></script></Td>
					<%End If%>
				</tr>
			</table>
		</td></Tr>
	</Table>
	<%If strPage<>"" Then%>
	<%Response.Write(XjLine(2,web_info(8),""))%>
	<Table class=xtable width="<%=web_info(8)%>" cellpadding=0 cellspacing=0 border=0>
		<Tr><Td  class=ctd height=25>
			<Table border=0 cellpadding=5 cellspacing=0 width="100%" height="100%">
				<tr><td align=Right width=*><%=pageLink(strPage)%></td></tr>
			</Table>
		</td></Tr>
	</Table>
	<%End If%>
	 <%Response.Write(XjLine(2,web_info(8),""))%>
<%
End Function

Function BottomTable()
%>
	<%Response.Write(XjLine(2,web_info(8),""))%>
	<Table class=xtable width="<%=web_info(8)%>" cellpadding=0 cellspacing=0 border=0>
		<Tr><Td  class=ctd height=22>
			<Table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
				<tr><td align=center width=*><%=bottommenu%></td></tr>
			</table>
		</td></Tr>
	</table>
	<%Response.Write(XjLine(2,web_info(8),""))%>
	<Table class=xtable width="<%=web_info(8)%>" cellpadding=0 cellspacing=0 border=0>
		<Tr><Td  class=ctd height=22>
			<Table border=0 cellpadding=0 cellspacing=0 width="100%" height="100%">
				<tr><td align=center width=*>
					版权所有&copy;:<%=web_info(13)%> &nbsp;&nbsp;
					页面执行时间: <%=Round(((Timer()-StartTime)*1000),2)%> 毫秒 &nbsp;&nbsp;
					Version: <%=web_info(1)%>&nbsp;&nbsp;
					你的IP：<%=xjweb.userip(0)%>
				</td></tr>
			</table>
		</td></Tr>
		<Tr><Td height=3 class=td_frame></td></Tr>
	</table>
	<Script language="javascript">
		document.all.loading.style.visibility='hidden';
	</script>
<%
End Function

Function xjLine(iHeight, iWidth, xColor)
	strCode="<Table cellspacing=0 border=0 cellpadding=0 width="""&iWidth&"""><tr><td"
		If xColor="class" Then
			strCode=strCode & " class=td_frame "
		ElseIf xColor<>"" Then
			strCode=strCode & " style=background-color:" & xColor & "; "
		End If
		strCode=strCode & " height="&iHeight&"></td>" &_
			vbcrlf &  "</tr></table>"
		xjLine = strCode
End Function

Function TbTopic(info)
%>
	<Table cellspacing=0 border=0 cellpadding=2 width="100%">
		<tr><td height=4></td></tr>
		<tr><td align=center valign=middle>
		<font style=font-size:16;font-weight:bold;><%= info %></font>
		</td></tr>
		<tr><td height=2></td></tr>
	</Table>
<%
End Function

Function XjDate(dt, iKind)
	If Not IsDate(dt) Then  xjDate="&nbsp; " : Exit Function
	If Not isNumeric(iKind) Then iKind=1
	Select Case iKind
		Case 1	'2005年1月1日
			XjDate=Year(dt) & "年" & Month(dt) & "月" & Day(dt) & "日"
		Case 2	'2005年1月1日星期六
			XjDate=Year(dt) & "年" & Month(dt) & "月" & Day(dt) & "日 星期"
				Select Case (Weekday(dt))
					Case 1
						XjDate=XjDate & "日"
					Case 2
						XjDate=XjDate & "一"
					Case 3
						XjDate=XjDate & "二"
					Case 4
						XjDate=XjDate & "三"
					Case 5
						XjDate=XjDate & "四"
					Case 6
						XjDate=XjDate & "五"
					Case 7
						XjDate=XjDate & "六"
				End Select
		Case 3	'2005-1-1
			XjDate=Year(dt) & "-" & Month(dt) & "-" & Day(dt)
		Case Else
			XjDate=dt
	End Select
End Function


Rem Functions模块
Function FileInc(kind, str)
	If Not isNumeric(kind) Then ErrAdd("File_Inc -- Kind Wrong!") : Exit Function
	kind=CInt(kind)		'0--JS  1--Css
	Select Case kind
		Case 0
			If JsFiles<>"" Then
				JsFiles=JsFiles & "||" & str
			Else
				JsFiles=str
			End If
		Case 1
			If CssFiles<>"" Then
				CssFiles=CssFiles & "||" & str
			Else
				CssFiles=str
			End If
		Case Else
			ErrAdd("Include file " & str & " 出错!")
	End Select
End Function

Function ErrAdd(errstr)
	If ErrInfo<>"" Then ErrInfo=ErrInfo & "$$$" & errstr : Exit Function
	ErrInfo=errstr
End Function

Sub Rw(str)
	If isNull(str) Then str=""
	str=CStr(str)
	Response.Write(str)
End Sub

Function closeObj()
	Erase web_info
	If isObject(xjweb) Then Set xjweb=Nothing
	If isObject(Rs) Then Set Rs=Nothing
	If isObject(conn) Then Set conn=Nothing
End Function

'-------------------------------提示信息相关函数------------------------------------------------------------
Function GotoPrompt()
	Session("InfoCode")=infoCode
	Session("InfoTitle")=infoTitle
	Session("InfoContens")=infoContents
	Session("InfoPreUrl")=infoPreUrl
	Session("InfoNewUrl")=infoNewUrl
	infoCode="" : infoTitle="" : infoContents="" : infoPreUrl="" : infoNewUrl=""		'释放这些变量
	Response.Clear
	Call closeObj()
	Response.Redirect(web_info(2) & "prompt.asp")
End Function
'---------------------------------检测信息是否包含非法字符-----------------------------
Function ChkStr(sn_var)
	ChkStr=False
	If sn_var="" Or Len(sn_var)>20 Or InStr(sn_var,"|")>0 Or InStr(sn_var,":")>0 Or InStr(sn_var,"'")>0 Or InStr(sn_var,"""")>0 Or InStr(sn_var,chr(9))>0 Or InStr(sn_var,chr(10))>0 Or InStr(sn_var,chr(13))>0 Or InStr(sn_var,chr(32))>0 Then
		Exit Function
	End If
	ChkStr=True
End Function

'---------------------------------处理相关内容(空值\数字,日期等)------------------------------
Function Var_Null(ub)
	Var_Null=Trim(ub)
	If Var_Null="" Or IsNull(Var_Null) Then Var_Null=""
End Function

Function Int_True(nvar)
	Int_True=True
	If Var_Null(nvar)="" Or Not(IsNumeric(nvar)) Or InStr(nvar,".")>0 Then Int_True=False
End Function

Function Num_True(nvar)
	Num_True=True
	If Var_Null(nvar)="" Or Not(IsNumeric(nvar)) Then Num_True=False
End Function

Function NullToStr(str)
	If IsNull(str) Then NullTostr=str : Exit Function
	NullTostr=str
End Function

Function NullToNum(inum)
	If IsNull(inum) Or inum="" Or Not(IsNumeric(inum)) Then NullToNum=0 : Exit Function
	NullToNum=inum
End Function
'------------------------------检测权限-----------------------------------------------------
Function ChkAble(str)
	'str 为权限位的位数,如第四位则str=4, 如果同时有多种可能性可用逗号(",")隔开 如: chkable(4,5,6)
	'注意: 如果str中含有-1则表示所有用户(包括客人)均有权限. 含有0 则表示所有登录用户具有权限
	ChkAble=False
	'If IsDebug Then ChkAble=True : Exit Function
	Dim tmpInt, tmpNum
	tmpInt=Split(str,",")
	for i=0 to ubound(tmpInt)
		tmpNum=tmpInt(i)
		If Not IsNumeric(tmpNum) Then tmpNum=0
		tmpNum=CInt(tmpNum)			'此句让num由字符变为数字
		If tmpNum=-1 Then ChkAble=True : Exit Function	'当str=-1时 表示所有用户(包括客人)均有权限
		If tmpNum=0 And Not IsNull(Session("userName")) And Session("userName")<>"" Then ChkAble=True : Exit Function	'当str=0时 表示所有登录用户具有权限
		If tmpNum>Len(Session("userAble")) Then tmpNum=Len(Session("userAble"))
		If tmpNum>0 Then
			If Mid(Session("userAble"),tmpNum,1)=1 Then ChkAble=True : Exit For
		End If
	Next
End Function

Function ChkPageAble(str)
	If Not(ChkAble(str)) Then
		infoCode=""
		infoTitle="没有权限"
		infoContents="所需权限:" & str & "<br>" &_
			"权限说明:<br>"&web_info(11)&"<br>"
		If IsNull(Session("userName")) Or Session("userName")="" Then
			infoContents=infoContents & "<li>可能是您还没有登录,请先<a href=login.asp>登录</a><br>"
			infoContents=infoContents & "<li>如果有其他疑问请<a href=mailto:zul@chinatrinity.com >联系系统管理员!</a>"
		End If
		infoPreUrl=""
		infoNewUrl=""
		Call GotoPrompt()
		Response.End
	End If
End Function

Function ChkAdminAble()
	If Not(ChkAble(1)) Or Session("admin")="" Then
		infoCode=""
		infoTitle="没有权限"
		infoContents="所需权限:系统管理员并进行后台登录<br>"
		If IsNull(Session("userName")) Or Session("userName")="" Then
			infoContents=infoContents & "<li>可能是您还没有登录,请先<a href=login.asp>登录</a><br>"
			infoContents=infoContents & "<li>如果有其他疑问请<a href=mailto:zul@chinatrinity.com >联系系统管理员!</a>"
		End If
		infoPreUrl=""
		infoNewUrl=""
		Call GotoPrompt()
		Response.End
	End If
End Function

Function ChkDepart(str)
	If not Session("userdepart")=str Then
		infoCode=""
		infoTitle="没有权限"
		infoContents="所需权限:必须是<b>" & str & "</b>成员<br><li>如果有其他疑问请<a href=mailto:zul@chinatrinity.com >联系系统管理员!</a>"
		infoPreUrl=""
		infoNewUrl=""
		Call GotoPrompt()
		Response.End
	End If
End Function
'--------------------------------------自动跳转页面代码---------------------------------------------
Function AutoRefresh(tTime)			'tTime为自动跳转的时间
	Dim prePage, prePageInfo
	prePage=LCase(Request("preUrl"))
	prePageInfo="前一页面"
	If prePage="" Then prePage=Request.ServerVariables("HTTP_REFERER")
	If Instr(prePage,"log") > 1 Or Instr(prePage,"prompt") > 1 Then prePage ="index.asp" : prepageinfo="首页"
	strCode="<span id=""downclock"" name=""downclock"">"&tTime&"</span> 秒后自动跳转到: "&_
		"<a href="""&prepage&""">"& prepageinfo &"</a>" &_
		vbcrlf & "<meta http-equiv=""refresh"" content="""& tTime &";url="&prepage&""">" &_
		vbcrlf & "<script language=""javascript"">" &_
		vbcrlf & "var totaltime = "&tTime&";	//倒计时秒数" &_
		vbcrlf & "function countDown()" &_
		vbcrlf & "{" &_
		vbcrlf & "downclock.innerHTML = totaltime;" &_
		vbcrlf & "window.setTimeout('countDown();',1000);"&_
		vbcrlf & "totaltime -= 1;"&_
		vbcrlf & "}" &_
		vbcrlf & "window.setTimeout('countDown();',1);"&_
		vbcrlf & "</script> "
	AutoRefresh=strCode
End Function

Function JsAlert(str,url)
	closeObj()
%>
	<Script language="javascript">
		alert("<%=str%>");
		<%If Trim(url)<>"" Then%>
			location.href="<%=Trim(url)%>";
		<%Else%>
		history.go(-1);

		<%End If%>
	</Script>

<%
	Response.End
End Function


Function JsPrompt(str)
	closeObj()
%>
	<Script language="javascript">
		alert("<%=str%>");
		window.close();
	</Script>
<%
	Response.End
End Function

Function FastLogin()
	If Session("userName")<>"" Then
		Rw("<font style=""font-size:16px;font-weight:bold;"">" & Session("userName") & "</font> : 您好!<br><br>您已经登录系统!<br>系统欢迎您!")
	Else
		Dim userName, userPwd, SaveTime
		userName=Request.Cookies(web_info(10))("userName")
		userPwd=Request.Cookies(web_info(10))("userPwd")
		SaveTime=Request.Cookies(web_info(10))("saveTime")
		If Not IsNumeric(SaveTime) Then SaveTime=0
%>
		<Table cellpadding=2 cellspacing=0 border=0 width=160>
			<Form name="frm_login" action="login.asp?action=login" method="post" onsubmit='return login_true();'>
			<Tr><Td align=Right>用户名称:</Td>
			<Td><input type="text" name=userName size=12 value="<%=userName%>" style="background-image:url(images/login_bg.gif);background-position:right;background-repeat:no-repeat;"></Td></Tr>
			<Tr><Td align=Right>用户密码:</Td>
			<Td>
				<input type="password" name=userPwd size=12 value="<%=userPwd%>">
			</Td></Tr>
			<Tr><Td align=Right>保存时间:</Td><Td>
				<Select name="SaveTime">
					<Option value=0 <%If SaveTime=0 Then Response.Write("Selected")%>>不保存</Option>
					<Option value=1 <%If SaveTime=1 Then Response.Write("Selected")%>>保存一天</Option>
					<Option value=31 <%If SaveTime=31 Then Response.Write("Selected")%>>保存一个月</Option>
					<Option value=365 <%If SaveTime=365 Then Response.Write("Selected")%>>保存一年</Option>
				</Select>
			</Td></Tr>
			<Tr><Td colspan=2 align=center><input type="submit" value=" 登 录 "></Td></Tr>
			</Form>
		</Table>
<%
	End If
End Function

'----------------------------------查找流水号的代码---------------------------------------------------
Function SearchLsh()
%>
	<Table border=0 cellpadding=2 cellspacing=0 width="100%">
		<Form name=frm_searchlsh action="<%=Request.Servervariables("SCRIPT_NAME")%>" method=post onsubmit='return searchlsh_true();'>
			<tr><td>
				&nbsp;&nbsp;输入流水号:
				<input tabindex=1 type=text name=s_lsh size=15 value="<%=Trim(Request("s_lsh"))%>">
				<input type="submit" value=" 查 找 ">
			</td></tr>
			</Form>
	</Table>
<%
End Function

'----------------------------------查找修理单号号的代码---------------------------------------------------
Function Searchxldh()
%>
	<Table border=0 cellpadding=2 cellspacing=0 width="100%">
		<Form name=frm_searchxldh ction="<%=Request.Servervariables("SCRIPT_NAME")%>" method=post onsubmit='return searchxldh_true();'>
			<tr><td>
				&nbsp;&nbsp;输入修理单号:
				<input tabindex=1 type=text name=s_xldh size=15 value="<%=Trim(Request("s_xldh"))%>">
				<input type="submit" value=" 查 找 ">
			</td></tr>
			</Form>
	</Table>
<%
End Function

'---------------------------------------发送站内短信--------------------------------------------------------
Function SendMsg(incept, sender, title, content)
	strSql="insert into ims_message (incept, sender, title, content) values ('"&incept&"', '"&sender&"', '"&title&"', '"&content&"')"
	Call xjweb.Exec(strSql, 0)
	If isdebug Then response.write("<script language=""javascript"">alert('短信发送成功!')</script>")
End Function


'----------------------------文件操作(FSO)----------------------------------------
Function Code_Fso(fString,ft1,ft2)
	Dim strTemp
	strTemp=Trim(fString)
	If strTemp="" Or IsNull(strTemp) Then : Code_Fso="" : Exit Function
	strTemp=Replace(strTemp,"""","\""")
	If ft2=1 Then
		strTemp=Replace(strTemp,":","")
		strTemp=Replace(strTemp,"|","")
	End If
	Select Case ft1
		Case 1
			strTemp=Replace(strTemp,vbcrlf,"<br>")
	End Select
	Code_Fso=strTemp
End Function

Function File_trim_vbcrlf(fvar)
	Dim temp1,tmp,tmpvar
	temp1=fvar
	tmp=False
	Do While Not tmp
		tmpvar=Left(temp1,1)
		If tmpvar=chr(10) or tmpvar=chr(13) Then
			temp1=Right(temp1,Len(temp1)-1)
		Else
			tmp=True
		End If
	Loop
	tmp=false
	Do While Not tmp
		tmpvar=Right(temp1,1)
		If tmpvar=chr(10) or tmpvar=chr(13) Then
			temp1=left(temp1,Len(temp1)-1)
		Else
			tmp=True
		End If
	Loop
  file_trim_vbcrlf=temp1
End Function

Sub Del_File(fname,ftype)
	'on error resume next
	Dim fobj,file_name,upload_path
	If Len(fname)<3 Then Exit Sub
	If Int(InStr(fname,"://"))>0 Then Exit Sub
	upload_path=web_Dim(13)
	If Right(upload_path,1)<>"/" Then upload_path=upload_path&"/"
	Select Case ftype
		Case 0
			file_name="style/"&fname
		Case 1
			upload_path=web_Dim(13)
			If Right(upload_path,1)<>"/" Then upload_path=upload_path&"/"
			file_name=upload_path&fname
		Case 5
			file_name=fname
		Case Else
			Exit Sub
	End Select
	file_name=Server.MapPath(file_name)
	Set fobj=CreateObject("Scripting.Filesystemobject")
	If fobj.fileexists(file_name) Then
		fobj.deletefile(file_name)
	End If
	Set fobj=Nothing
End Sub

Function get_file(file_name)
	Dim filetemp,fileos,filepath
	Set fileos=CreateObject("Scripting.Filesystemobject")
	filepath=Server.MapPath(file_name)
	Set filetemp=fileos.opentextfile(filepath,1,True)
	get_file=filetemp.ReadAll
	filetemp.close
	Set filetemp=Nothing
	Set fileos=Nothing
End Function

Sub create_file(file_name,filetype)
	Dim filetemp,fileos,filepath
	Set fileos=CreateObject("Scripting.FileSystemObject")
	filepath=Server.MapPath(file_name)
	Set filetemp=fileos.createtextfile(filepath,True)
	filetemp.writeline(filetype)
	filetemp.close
	Set filetemp=Nothing
	Set fileos=Nothing
End Sub

'---------------------------------检测是否启用Cookies----------------------------
Function CheckCookies()
	Dim strCookies
	Dim tempSec
	tempSec=3	'刷新页面时间
	strCookies=Request.Cookies("enablecookies")
	If IsNull(strCookies) Or strCookies="" Or strCookies<>"enable" Then
		Response.Cookies("enablecookies")="enable"
		Response.Cookies("enablecookies").Expires=date+3650
		Dim strTemp
		strTemp="<html>" &_
			vbcrlf & "<head>" &_
			vbcrlf & "<title>检测您的浏览器是否启用Cookies</title>" &_
			vbcrlf & "<meta http-equiv=""refresh"" content="""&tempsec&""">" &_
			vbcrlf & "<link href="""&web_info(2)&"styles/styles.css"" rel=""stylesheet"" type=""text/css"">" &_
			vbcrlf & Comm_Css &_
			vbcrlf & "</head>" &_
			vbcrlf & "<body>" &_
			vbcrlf & "<table height=""80%"" width=""100%"" border=""0"" style=""text-align:center;""><tr><td>" &_
			vbcrlf & "<table cellspacing=0 class=xtable height=220 width=450><tr><th class=th height=30>" &_
			vbcrlf & "系统检测您的浏览器是否开启Cookies" &_
			vbcrlf & "</td></tr><tr><td class=ltd>" & _
			vbcrlf & "<ul>" &_
			vbcrlf & "<li>欢迎使用本系统,本系统需要使用Cookies" &_
			vbcrlf & "<li>如果您的浏览器不支持Cookies本系统将无法正常运行" &_
			vbcrlf & "<li>正在检测Cookies<span id='dottt'>.</span>" &_
			vbcrlf & "<li><span id='clock'>"&tempsec&"</span> 秒后自动跳转" &_
			vbcrlf & "</ul><ul><b>您的系统信息</b>" &_
			vbcrlf & "<li>您的操作系统:" & xjweb.usersys(0) &_
			vbcrlf & "<li>您的浏 览 器:" & xjweb.usersys(1) &_
			vbcrlf & "<li>显示器分辨率:<script language=""javascript"">document.write(screen.width + '×' + screen.height)</script>" &_
			vbcrlf & "<li>您的  真实 IP:" & xjweb.userip(0) &_
			vbcrlf & "</ul>" &_
			vbcrlf & "</td></tr></table>" &_
			vbcrlf & "</td></tr></table>" &_
			vbcrlf & "<script language=""JavaScript"">" &_
			vbcrlf & "<!--" &_
			vbcrlf & "var totaltime = "&tempsec&";	//倒计时秒数" &_
			vbcrlf & "var tstr = '..';" &_
			vbcrlf & "function countdown()" &_
			vbcrlf & "{" &_
			vbcrlf & "clock.innerHTML = totaltime;" &_
			vbcrlf & "dottt.innerHTML=tstr" &_
			vbcrlf & "window.setTimeout('countdown()',1000);" &_
			vbcrlf & "totaltime -= 1;" &_
			vbcrlf & "tstr=tstr+'.';" &_
			vbcrlf & "}" &_
			vbcrlf & "//-->" &_
			vbcrlf & "window.setTimeout('countdown()',1);" &_
			vbcrlf & "</script>" &_
			vbcrlf & "</body>" &_
			vbcrlf & "</html>"
		Response.Write strTemp
		Response.End
	End If
End Function

Function SiteStat()
	Dim userIP, strAgent,strUser
	userIP=xjweb.userip(0)
	strAgent=Trim(LCase(Request.ServerVariables("HTTP_USER_AGENT")))
	strUser=Session("userName")
	If IsNull(strUser) Or strUser="" Then strUser="客人" & replace(userIP,".","")

	Rem 在线用户统计
	'---------------------------删除过时的数据信息---------------------------------------
	strSql="delete from [ims_online] where datediff('n',ol_lasttime,'"&now()&"')>60"  'datediff("n", var1,var2)---n为分钟 60不刷新删除
	Call xjweb.Exec(strSql, 0)
	'---------------------------记录访问系统的用户信息--------------------------------
	strSql="select * from [ims_online] where ol_ip='"&userIP&"'"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.Eof Or Rs.Bof Then
		strSql="insert into [ims_online] (ol_user, ol_ip, ol_logintime, ol_lasttime,ol_onurl, ol_agent) values ('"&strUser&"','"&userIP&"','"&now()&"','"&now()&"','"&CurPage&"','""')"
		Call xjweb.Exec(strSql, 0)
		'进行访问用户统计
		strSql="insert into [ims_stat] (stat_user, stat_ip, stat_agency, stat_time) values ('"&strUser&"','"&userIP&"','"&strAgent&"','"&Now()&"')"
		strSql="delete from [ims_stat] where datediff('m',stat_time,'"&now()&"')>3"  '删除3个月以前的的访问统计
		Call xjweb.Exec(strSql, 0)
	Else
		strSql="update [ims_online] set ol_user='"&strUser&"', ol_lasttime='"&now()&"', ol_onurl='"&CurPage&"' where ol_ip='"&userIP&"'"
		Call xjweb.Exec(strSql, 0)
	End If
	Rs.Close
End Function
%>