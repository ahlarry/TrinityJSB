<!--#include file="include/conn.asp"-->
<%
CurPage="模具质量"
strPage="mquality"
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
			Welcome to 模具质量模块
		</Td></Tr>
	</Table>
<%
End Sub
%>