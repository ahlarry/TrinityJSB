<!--#include file="include/conn.asp"-->
<%
CurPage="分值统计"
strPage="mtstat"
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
			Welcome to 分值统计模块
		</Td></Tr>
	</Table>
<%
End Sub
%>