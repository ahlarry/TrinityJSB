<!--#include file="include/conn.asp"-->
<%
Dim action
CurPage="�鿴�������"
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

Rem ����Ϊ��ҳ����
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
	If Rs.Eof Or Rs.Bof Then Rw("��ǰֻ��������!") : Exit Function
	%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="<%=web_info(8)%>">
		<Tr>
			<Td class=th>ID</Td>
			<Td class=th>����</Td>
			<Td class=th>IP</Td>
			<Td class=th>��¼ʱ��</Td>
			<Td class=th>����ʱ��</Td>
			<Td class=th>����ҳ��</Td>
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