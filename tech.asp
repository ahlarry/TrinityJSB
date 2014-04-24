<!--#include file="include/conn.asp"-->
<%
CurPage="问题分析"
strPage="tech"
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
			<%Call TbTopic("Welcome to 问题分析")%>
		</Td></Tr>
	</Table>
<%
End Sub
%>