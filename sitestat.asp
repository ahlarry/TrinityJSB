<!--#include file="include/conn.asp"-->
<%
CurPage="ϵͳͳ��"
strPage=""	 'sitestat
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
			Welcome to ϵͳͳ��ģ��
		</Td></Tr>
	</Table>
<%
End Sub
%>