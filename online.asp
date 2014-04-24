<!--#include file="include/conn.asp"-->
<%
Dim action
CurPage="查看在线情况"
action=LCase(request("action"))
Select Case action
	Case "count"
		Call OnlineCount()
	Case "list"
		Call xjweb.Header()
		Call TopTable()
		Call OnlineList()
		Call BottomTable()
		Call xjweb.Footer()
		Call CloseObj()
	Case Else
		Call Main()
End Select
closeObj()

Rem 下面为本页函数
Sub Main()
	xjweb.header()
	rw(TopTable())
	'Call Main()
	rw(BottomTable())
	xjweb.footer()
	Call closeObj()
End Sub
Function OnlineCount()
	rw("document.write(" & xjweb.RsCount("ims_online") & ")")
End Function

Function OnlineList()
	Set Rs=xjweb.Exec("select * from [ims_online]",1)
	If Rs.Eof Or Rs.Bof Then Rw("当前只有你在线!") : Exit Function
	%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="<%=web_info(8)%>">
		<Tr>
			<Td class=th>ID</Td>
			<Td class=th>名称</Td>
			<Td class=th>IP</Td>
			<Td class=th>登录时间</Td>
			<Td class=th>在线时长</Td>
			<Td class=th>所在页面</Td>
		</Tr>
	<%
	i=1
	Do While Not Rs.Eof
	%>
		<Tr>
			<Td class=ctd><%=i%></Td>
			<Td class=ctd><%=Rs("ol_user")%></Td>
			<Td class=ctd><%=Rs("ol_ip")%></Td>
			<Td class=ctd><%=Rs("ol_logintime")%></Td>
			<Td class=ctd><%=datediff("n",Rs("ol_logintime"),now())%></Td>
			<Td class=ctd><%=Rs("ol_onurl")%></Td>
		</Tr>
	<%
		i=i+1
		Rs.MoveNext
	Loop
	%>
	</Table>
	<%
	Rs.Close
End Function
%>