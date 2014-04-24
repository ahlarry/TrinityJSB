<!--#include file="include/conn.asp"-->
<%
CurPage="短信息"
strPage=""	 'aboutus
xjweb.header()
	dim iid, struname, bnew, action
	iid = request("id")
	struname=session("userName")
	bnew=request("new")
	action=request("action")
	if not(isnumeric(iid)) then iid=0

	if bnew="" then bnew=false

	Call Main()
xjweb.footer()
closeObj()

Sub Main()
	select case action
	case "send"
		call sendboxmsg()
	case "incept"
		call inceptboxmsg()
	case else
		call inceptboxmsg()
	end select
End Sub

function inceptboxmsg()
	if bnew then	'是未读信息
		call displaymsg(0)
	else		'全部信息
		call displaymsg(1)
	end if
end function

function displaymsg(chk)
	if chk=0 then
		if iid<>0 then
			strSql="select * from [ims_message] where id="&iid&" and incept='"&struname&"' and delr=0 order by id"
		else
			strSql="select * from [ims_message] where incept='"&struname&"' and delr=0 order by id"
		end if
	elseif chk=1 then
		if iid<>0 then
			strSql="select * from [ims_message] where id="&iid&" and incept='"&struname&"' and delr<>2 order by id"
		else
			strSql="select * from [ims_message] where incept='"&struname&"' and delr<>2 order by id"
		end if
	end if
	set rs=xjweb.exec(strSql,1)
	if rs.eof or rs.bof then
		response.write("暂时没有未读短信息!")
	else
	%>
<table border="0" cellspacing="0" cellpadding="3" class=xtable width="90%">
  <caption>
  <b><%=session("userName")%>的短信息</b>
  </caption>
  <tr>
    <td class=rtd width=100>标题:</td>
    <td class=ltd width="*"><%=rs("title")%></td>
  </tr>
  <tr height=200>
    <td class=rtd valign="center">内容:</td>
    <td class=ltd valign="top"><%=rs("content")%></td>
  </tr>
  <tr>
    <td class=rtd>发送:</td>
    <td class=ltd><%=rs("sender")%>于<%=rs("sendtime")%>发送</td>
  </tr>
  <%if rs("delr")=0 then%>
  <form action="msg_indb.asp?action=affirm" method="post">
    <tr>
      <td colspan=2 class=ctd alt="点击收到确认,将不再弹出此信息!"><a href='uctrl_sendmsg.asp?title=Re:<%=rs("title")%>&incept=<%=rs("sender")%>' target=_blank>回  复</a>&nbsp;&nbsp;
        <input type="submit" value=" 收到确认 ">
        &nbsp;&nbsp;<a href="msg_indb.asp?id=<%=rs("id")%>&action=Sdelete">#删 除#</td>
    </tr>
    <input type="hidden" name="id" value="<%=rs("id")%>">
  </form>
  <%end if%>
</table>
<%
		response.write prenext(rs, chk)
	end if
	rs.close
	strSql="update [ims_message] set flag=1 where id="&clng(iid)&""
	Call xjweb.Exec(strSql, 0)
end function

function prenext(rs, ichk)
	prenext="<table border=""0"" width=""90%"" cellpadding=""4"" cellspacing=""0""><tr><td width=""50%"" align=""left"">&nbsp;"
	dim prs, strtitle
	if ichk=0 then
		strSql="select id,title from ims_message where id<" & rs("id") &" and incept='"&rs("incept")&"' and delr=0 order by id desc"
	else
		strSql="select id,title from ims_message where id<" & rs("id") &" and incept='"&rs("incept")&"' and delr<2 order by id desc"
	end if
	set prs=xjweb.Exec(strSql, 1)
	if not(prs.eof or prs.bof) then
		if len(prs("title")) > 8 then strtitle=left(prs("title"),8) & "......" else strtitle=prs("title") end if
		prenext=prenext & "上一封:<a href="""&request.servervariables("script_name")&"?"
		if ichk=0 then prenext=prenext & "new=true&"
		prenext=prenext &"id="&prs("id")&""">"&strtitle&"</a>"
	end if
	prs.close
	prenext=prenext & "</td><td width=""*"" align=""right"">"
	if ichk=0 then
		strSql="select id,title from [ims_message] where id>" & rs("id") &" and incept='"&rs("incept")&"' and delr=0 order by id"
	else
		strSql="select id,title from [ims_message] where id>" & rs("id") &" and incept='"&rs("incept")&"' and delr<2 order by id"
	end if
	set prs=xjweb.Exec(strSql, 1)
	if not(prs.eof or prs.bof) then
		if len(prs("title")) > 8 then strtitle=left(prs("title"),8) & "......" else strtitle=prs("title") end if
		prenext=prenext & "下一封:<a href="""&request.servervariables("script_name")&"?"
		if ichk=0 then prenext=prenext & "new=true&"
		prenext=prenext &"id="&prs("id")&""">"&strtitle&"</a>"
	end if
	prs.close
	set prs=nothing
	prenext=prenext & "&nbsp;</td></tr></table>"
end function

function sendboxmsg()
	if iid<>0 then
		strSql="select * from ims_message where id="&iid&" and sender='"&struname&"' and dels<>2 order by id"
	else
		strSql="select * from ims_message where sender='"&struname&"' and dels<>2 order by id"
	end if
	set rs=xjweb.Exec(strSql,1)
	if rs.eof or rs.bof then
		response.write("发件箱中暂时没有短信息!")
	else
	%>
<table border="0" cellspacing="0" cellpadding="3" class=xtable width="90%">
  <caption>
  <b><%=session("userName")%>发送的短信息</b>
  </caption>
  <tr>
    <td class=th width=80>标题</td>
    <td class=ltd width="*"><%=rs("title")%></td>
  </tr>
  <tr height=200>
    <td class=th valign="center">内容</td>
    <td class=ltd valign="top"><%=rs("content")%></td>
  </tr>
  <tr>
    <td class=th>接收</td>
    <td class=ltd><%=rs("incept")%></td>
  </tr>
  <tr>
    <td class=th>发送时间</td>
    <td class=ltd><%=rs("sendtime")%></td>
  </tr>
</table>
<%
		response.write sendprenext(rs)
	end if
	rs.close
end function

function sendprenext(rs)
	sendprenext="<table border=""0"" width=""90%"" cellpadding=""4"" cellspacing=""0""><tr><td width=""50%"" align=""left"">&nbsp;"
	dim prs, strtitle

	strSql="select id,title from [ims_message] where id<" & rs("id") &" and sender='"&rs("sender")&"' and dels<>2 order by id desc"

	set prs=xjweb.Exec(strSql, 1)
	if not(prs.eof or prs.bof) then
		if len(prs("title")) > 8 then strtitle=left(prs("title"),8) & "......" else strtitle=prs("title") end if
		sendprenext=sendprenext & "上一封:<a href="""&request.servervariables("script_name")&"?action=send&id="&prs("id")&""">"&strtitle&"</a>"
	end if
	prs.close
	sendprenext=sendprenext & "</td><td width=""*"" align=""right"">"

	strSql="select id,title from [ims_message] where id>" & rs("id") &" and sender='"&rs("sender")&"' and dels<>2 order by id"

	set prs=xjweb.Exec(strSql, 1)
	if not(prs.eof or prs.bof) then
		if len(prs("title")) > 8 then strtitle=left(prs("title"),8) & "......" else strtitle=prs("title") end if
		sendprenext=sendprenext & "下一封:<a href="""&request.servervariables("script_name")&"?action=send&id="&prs("id")&""">"&strtitle&"</a>"
	end if
	prs.close
	set prs=nothing
	sendprenext=sendprenext & "&nbsp;</td></tr></table>"
end function
%>
<body onUnload="opener.location.reload()">
