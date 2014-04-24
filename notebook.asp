<!--#include file="include/conn.asp"-->
<%
CurPage="系统留言"
strPage="notebook"
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=4 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%Call NoteBook()%>
		</Td></Tr>
	</Table>
<%
End Sub

Function NoteBook()
	Call TbTopic(web_info(0) & "系统留言")
%>
	<table border="0" width="95%" cellpadding="2" cellspacing="0">
	<tr>
		<td>
		<%
			Dim PerPageCount, absPageNum, iNum, absRecordNum
			PerPageCount = 20
			Set Rs = Server.CreateObject("ADODB.RECORDSET")
			Rs.cachesize = PerPageCount
			strSql="select * from [notebook] order by id desc"
			Call xjweb.Exec("", -1)
			Rs.open strSql,conn,1,3
			Rs.pagesize = PerPageCount
			absPageNum = request("ipage")
			If absPageNum = "" Then absPageNum = 1
			If not(isnumeric(absPageNum)) Then absPageNum = 1
			absPageNum = clng(absPageNum)
			If absPageNum > Rs.pagecount Then absPageNum = Rs.pagecount

			Rs.absolutepage = absPageNum
			iNum = Rs.recordcount - (PerPageCount * (absPageNum - 1))

			for absRecordNum = 1 to PerPageCount
			%>
			<table width="100%" cellpadding=2 cellspacing=0 class=xtable align="center">
				<tr>
					<td class=ctd width="150" rowspan="4" valign="top">
					<%
						Dim tempusername, tempuserface, temprs, tempstr
						tempstr = Rs("username")
						strSql="select user_nick,user_face from ims_user where user_name = '"&tempstr&"'"
						Set temprs = xjweb.Exec(strSql, 1)
						If temprs.eof Then
							tempusername = Rs("username") & "<br>(此用户已删除)"
							tempuserface = "noface.gif"
						else
							If isnull(temprs("user_nick")) Then
								tempusername = Rs("username")
							else
								tempusername = temprs("user_nick")
							end If

							If isnull(temprs("user_face")) Then
								tempuserface = "noface.gif"
							else
								tempuserface = temprs("user_face")
							end If
						end If
					%>
						<img src="<%=web_info(2) & "images/face/" & tempuserface%>" onload="javascript:if(this.width>120) this.width=120;if(this.height>120) this.height=120;"></img><br><br>
						<%=tempusername%>
					</td>
					<td class=ltd width="*">留言时间: <%=Rs("indate")%></td>
					<td class=ctd width="15%">第 <b><%=iNum%></b> 条 留言</td>
				</tr>
				<tr>
					<td class=ltd colspan="2"><%=xjweb.htmltocode(Rs("content"))%>
					<%If not(isnull(Rs("editdate"))) Then%>
						<br><br><br>
						<div align="right">编缉时间: <%=Rs("editdate")%></div>
					<%end If%>
					</td>
				</tr>
				<%If not(isnull(Rs("reply"))) Then%>
					<tr>
						<td class=ltd colspan="2"><b>管理员回复:</b><br><%=xjweb.htmltocode(Rs("reply"))%>
						</td>
					</tr>
				<%end If%>
				<form action="notebook_indb.asp" method="post" onsubmit="return confirm('删除后将不能恢复!确信吗?');">
				<tr>
					<td class=rtd colspan="2">
						<%If chkable(1) Then%>
							<a href="notebook_reply.asp?id=<%=Rs("id")%>">回复</a>
						<%end If%>
						<%If session("userName") = Rs("username") Then%>
							<a href="notebook_change.asp?id=<%=Rs("id")%>">编缉</a>
						<%end If%>
						<%If chkable(1) Then%>
							<input type="submit" value=" 删除 ">
						<%end If%>
						&nbsp;
					</td>
				</tr>
				<input type="hidden" name="id" value=<%=Rs("id")%>>
				<input type="hidden" name="indbinf" value="delete">
				</form>
				</table>

				<table border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td height="5">
						</td>
					</tr>
				</table>
			<%
					iNum = iNum - 1
					Rs.movenext
					If Rs.eof Then
						exit for
					end If
				next
			%>
			<br>
			<table width="100%" border="0" cellpadding="3" cellspacing="0" class=xtable   align="center" onmouseover="this.bgcolor='#dddddd'" onmouseout="this.bgcolor='';">
				<tr>
					<td class=ctd width="*">
					总共 <b><%=Rs.recordcount%></b> 个留言  每页 <b><%=PerPageCount%></b> 个留言 <br>
						<%
							for i = 1 to Rs.pagecount
							'for i = 1 to 100
								If i = absPageNum Then
									response.write("<font style='font-size:10pt;font-weight:bold;'>第" & i & "页</font>")
								else
									response.write("<a href='?ipage="&i&"'>第" & i & "页</a>")
								end If
								response.write(" ")
								If i mod 12 = 0 Then response.write("<br>")
							next
						%>
					</td>
				</tr>
			</table><br>
		</td>
	</tr>
	</table>
<%
End Function
%>