<!--#include file="include/conn.asp"-->
<%
'Call ChkPageAble(2)
CurPage="首页"
strPage=""  'index
Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Rem 以下为本文件的函数部分
Sub Main()
%>
	<Table cellpadding=0 cellspacing=0 width="<%=web_info(8)%>">
		<tr>
		<td width=180 valign=top>
			<%Call UserLogin()%>
			<%Response.Write(xjLine(2,"100%",""))%>
			<%=TaskStat()%>
			<%Response.Write(xjLine(2,"100%",""))%>
			<%'=UserStat()%>
			<%'Response.Write(xjLine(2,"100%",""))%>
			<%=AuthorInfo()%>
		</td>
		<td width=2></td>
		<td width=* valign=top>
			<%'=IndexCenter()%>
			<%=DayReciteWords(10)%>
		</td>
		<td width=2></td>
		<td width=180 valign=top>
			<%=CurTask()%>
			<%Response.Write(xjLine(2,"100%",""))%>
			<%=CurTestMould()%>
			<%'Response.Write(xjLine(2,"100%",""))%>
			<%'=OnlineUser()%>
		</td>
		</tr>
	</table>
<%
End Sub
Function UserLogin()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> ■用户登录</td></tr>
		<Tr height=100><td class=ctd><%=FastLogin()%></td></tr>
	</Table>
<%
End Function
Function TaskStat()
	dim tmp
	tmp="任务总数：" & xjweb.rscount("mtask") & " 套$$$" &_
		"去年任务：" & xjweb.rscount("mtask where Year(rwxdsj)=(Year(now)-1)") & " 套$$$" &_
		"今年目前：" & xjweb.rscount("mtask where Year(rwxdsj)=Year(now)") &" 套$$$" &_
		"正在设计：" & xjweb.rscount("mtask where isnull(sjjssj)") &" 套"
	TaskStat=ListTop("任务书统计", "100%", "images\list\article_2.gif", tmp , "", 0)
End Function

Function AuthorInfo()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> ■作者信息</td></tr>
		<Tr><td class=ltd>
			<b>My MSN:</b><br>
				<span style=padding:10; alt="朱磊">ahLarry@hotmail.com<br></span>
			<b>My QQ:</b><br>
				<span style=padding:10; alt="朱磊">28521622<br></span>
			<b>My Email:</b><br>
				<a href=mailto:zul@chinatrinity.com ><span  style=padding:10; alt="这是我的Email，有事没事常联系！">zul@chinatrinity.com</span></a>
		</td></tr>
	</Table>
<%
End Function

Function UserStat()
	dim tmp
	tmp="用户总数：" & xjweb.rscount("[ims_user]") & " 人$$$" &_
		"组长：" & xjweb.rscount("[ims_user] where Mid(user_able,4,1)=1") &" 人$$$" &_
		"组员：" & xjweb.rscount("[ims_user] where Mid(user_able,5,1)=1") &" 人"
	UserStat=ListTop("用户统计", "100%", "images\list\dot_1.gif", tmp , "", 0)
End Function
Function CurTask()
	dim tmp
	tmp=""
	Set Rs=xjweb.Exec("Select top 5 lsh, dmmc, dwmc from [mtask] where isnull(sjjssj) order by lsh desc",1)
	Do While Not Rs.Eof
		if tmp<>"" Then
			tmp=tmp & "$$$<a href=mtask_list.asp?lsh="&rs("lsh")&" alt=""流水号:"&rs("lsh")&"<br>单位名称:"&rs("dwmc")&"<br>断面名称:"&rs("dmmc")&""">" & xjweb.StringCut(rs("dwmc") & "--" & rs("dmmc"), 12) & "</a>"
		else
			tmp="<a href=mtask_list.asp?lsh="&rs("lsh")&" alt=""流水号:"&rs("lsh")&"<br>单位名称:"&rs("dwmc")&"<br>断面名称:"&rs("dmmc")&""">" & xjweb.StringCut(rs("dwmc") & "--" & rs("dmmc"), 12) & "</a>"
		end if
		Rs.moveNext
	Loop
	Rs.Close
	CurTask=ListTop("正在设计任务", "100%", "images\list\article_1.gif", tmp , "mtask_list.asp?term=no", 0)
End Function
Function CurTestMould()
	dim tmp
	tmp=""
	Set Rs=xjweb.Exec("Select top 5 a.lsh as lsh, dmmc, dwmc from [mtask] a, [ts_mould] b where a.lsh=b.lsh and isnull(tsjssj) order by tsgxsj desc",1)
	i=1
	Do While Not Rs.Eof
		if tmp<>"" Then
			tmp=tmp & "$$$<a href=mtest_list.asp?lsh="&rs("lsh")&" alt=""流水号:"&rs("lsh")&"<br>单位名称:"&rs("dwmc")&"<br>断面名称:"&rs("dmmc")&""">" & xjweb.StringCut(rs("dwmc") & "--" & rs("dmmc"), 12) & "</a>"
		else
			tmp="<a href=mtest_list.asp?lsh="&rs("lsh")&" alt=""流水号:"&rs("lsh")&"<br>单位名称:"&rs("dwmc")&"<br>断面名称:"&rs("dmmc")&""">" & xjweb.StringCut(rs("dwmc") & "--" & rs("dmmc"), 12) & "</a>"
		end if
		Rs.moveNext
	Loop
	Rs.Close
	CurTestMould=ListTop("最新调试模具", "100%", "images\list\article_1.gif", tmp , "mtest_list.asp?term=no", 0)
End Function
Function OnlineUser()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;>
		■访问统计</td></tr>
		<Tr><td class=ltd>
			<li alt="自2005年2月15日以来">系统访问总数: <%=xjweb.rscount("[ims_stat]")%> 次</li>
			<li>当前在线用户: <a href="online.asp?action=list"><%=xjweb.rscount("[ims_online]")%></a> 人</li>
			<li>昨天访问次数: <%=xjweb.rscount("[ims_stat] where datediff('d',stat_time,'"&now()&"')=1")%> 次</li>
			<li>今天访问次数: <%=xjweb.rscount("[ims_stat] where datediff('d',stat_time,'"&now()&"')=0")%> 次</li>
		</td></tr>
	</Table>
<%
End Function

Function indexCenter()
	strSql="select user_name, user_birthday from [ims_user] where Month([user_birthday])=Month(now()) and Day([user_birthday])=Day(now()) order by user_name"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.Bof Or Rs.Eof Then
		Call indexMain()
	Else
		Call HappyBirthday(Rs)
	End If
	Rs.Close
End Function

Function indexMain()
%>
	<Table border=0 cellpadding=5 cellspacing=0 width="100%">
		<Tr><Td align=center valign=middle height=140>
			<div style='width:300;filter:glow(color=green,strength=5);font-size:18pt;color:white;'>2006挤模信息管理系统</div><br>
			<div style='width:300;filter:glow(color=blue,strength=5);font-size:11pt;color:white;'>Ver: <%=web_info(1)%></div><br>
			<div style='width:300;filter:glow(color=red,strength=3);font-size:14pt;color:white;'>欢迎使用</div>
		</Td></Tr>
	</Table>
<%
End Function

Function HappyBirthday(Rs)
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th> 生日祝福 </td></tr>
		<Tr><td class=ctd><img src="<%=web_info(2)%>images/birthday.jpg" onload="if(this.width>280) this.width=280;if(this.height>280) this.height=280;"><br>
		<%
			Do while not Rs.Eof
				Response.Write("<div alt=""出生日期:" &xjdate(rs("user_birthday"),1)&"""><font style=font-size:16pt;font-weight:bold;>" & Rs("user_name") & "</font> <font style=font-size:14pt;>生日快乐!</font></div>")
				Rs.moveNext
			Loop
		%>
		</td></tr>
	</Table>
<%
End Function

Function DayReciteWords(iCount)
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th colspan=4> 天天背单词 </td></tr>
		<%
		Dim Myvalue
		i=1
		Randomize '初始化随机数生成器。
		Myvalue = Int((6703 * Rnd) + 1)' 产生随机数。
		strSql="select * from [ims_words] where id > "& Myvalue &" and id <= "& Myvalue+iCount &" "
		Set Rs=xjweb.Exec(strSql,1)
		Do While Not Rs.Eof
		%>
			<Tr>
				<Td class=ctd>No.<b><%=i%></b></Td>
				<Td class=ctd><font color = #ff0000><b><%=Rs("word")%></b></font></Td>
				<Td class=ctd><font color=green><b>[<font face="Kingsoft Phonetic"><%=Rs("phonetic")%></font>]</b></font></Td>
				<Td class=ltd><font color=blue><%=Rs("Remark")%></font></Td>
			</Tr>
		<%
			i=i+1
			Rs.movenext
		Loop
		Rs.Close
		%>
	</Table>
<%
End Function

Function ListTop(topic, iwidth, listsymbol, content , infoMore, bDis)
	dim tmp, tmpCon, i
	tmp="<Table class=xtable cellpadding=2 cellspacing=0 width=""" & iwidth & """>" &_
		vbcrlf & "<Tr><td class=th style=text-align:left;> ■" & topic & "</td></tr>" &_
		vbcrlf & "<Tr><td class=ltd>" &_
		vbcrlf & "<Table cellpadding=3 cellspacing=0 width=""100%"">"

	tmpCon=split(content,"$$$")
	for i=0 to Ubound(tmpCon)
		tmp=tmp & vbcrlf & "<Tr><Td valign=top>"
		if listsymbol<>"" Then
			tmp=tmp & " <img src=""" & listsymbol & """  align=""absmiddle"" hspace=""3""> "
		else
			tmp=tmp & " · "
		end if
		tmp=tmp & tmpCon(i) &_
			vbcrlf & "</Td></Tr>" &_
			vbcrlf & "<Tr><Td height=1 background='images/bg_dian.gif' colspan='2'></Td></Tr>"
	next

	if infoMore<>"" Then
		tmp=tmp & vbcrlf & "<Tr><Td align=right>" &_
			vbcrlf & "<a href=""" & infoMore & """>more...</a>" &_
			vbcrlf & "</Td></Tr>"
	End if

	tmp=tmp & vbcrlf & "</Table>" &_
		vbcrlf & "</td></tr>" &_
		vbcrlf & "</Table>"

	if bDis Then
		Response.Write tmp
	else
		listTop=tmp
	end if
End Function
%>

