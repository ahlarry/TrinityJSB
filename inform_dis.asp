<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="系统通知 → 查看通知"
strPage="inform"
'Call FileInc(0, "js/docbak.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%Call InformDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function InformDisplay()
	Dim iid
	iid=request("id")
	if not(isnumeric(iid)) then iid=0
	iid=clng(iid)
	if iid=0 then
		strSql="select * from [ims_inform] order by id desc"
	else
		strSql="select * from [ims_inform] where id="&iid&""
	end if
	set rs=xjweb.Exec(strSql, 1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("暂时没有通知！","index.asp")
	Else
%>
	<table border="0" cellspacing="0" cellpadding="3" width="600" align="center">
		<tr>
			<td align="center">
				<%Call TbTopic(rs("inform_topic"))%>
				发布人:<%=rs("inform_lzr")%> 发布日期:<%=formatdatetime(rs("inform_date"),1)%> &nbsp;
				<%If chkable(1) then%>
					<a href="inform_change.asp?action=change&id=<%=rs("id")%>">更改</a>
					<a href="inform_change.asp?action=delete&id=<%=rs("id")%>" onclick="return confirm('您确认删除此通知吗?');">删除</a>
				<%end if%>
			</td>
		</tr>
		<tr><td><%response.write(xjLine(1, "100%", "class"))%></td></tr>
		<tr>
			<td height="200" valign="top" align="center">
				<Table cellpadding=4 cellspacing=0 width="95%">
				<Tr><Td>
				<%=Replace(Replace(rs("inform_content"),"  ","&nbsp;&nbsp;"),vbcrlf,"<br>")%>
				</Td></Tr>
				</Table>
			</td>
		</tr>
		<tr><td><%response.write(xjLine(1, "100%", "class"))%></td></tr>
	</table>
	<%=prenext(rs)%>
<%
	end if
	rs.close
End Function

function prenext(rs)
	prenext="<table border=""0"" width=""600"" cellpadding=""4"" cellspacing=""0""><tr><td width=""50%"" align=""center"">&nbsp;"
	dim prs, strtitle
	strSql="select id,inform_topic from [ims_inform] where id<" & rs("id") &" order by id desc"
	set prs=xjweb.Exec(strSql, 1)
	if not(prs.eof or prs.bof) then
		if len(prs("inform_topic")) > 8 then strtitle=left(prs("inform_topic"),8) & "......" else strtitle=prs("inform_topic") end if
		prenext=prenext & "Prev: <a href="""&request.servervariables("script_name")&"?"
		prenext=prenext &"id="&prs("id")&""">"&strtitle&"</a>"
	end if
	prs.close
	prenext=prenext & "</td><td width=""*"" align=""right"">"
	strSql="select id,inform_topic from [ims_inform] where id>" & rs("id") &" order by id"
	set prs=xjweb.Exec(strSql, 1)
	if not(prs.eof or prs.bof) then
		if len(prs("inform_topic")) > 8 then strtitle=left(prs("inform_topic"),8) & "......" else strtitle=prs("inform_topic") end if
		prenext=prenext & "Next: <a href="""&request.servervariables("script_name")&"?"
		prenext=prenext &"id="&prs("id")&""">"&strtitle&"</a>"
	end if
	prs.close
	set prs=nothing
	prenext=prenext & "&nbsp;</td></tr></table>"
end function
%>