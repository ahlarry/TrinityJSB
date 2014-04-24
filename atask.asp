<!--#include file="include/conn.asp"-->
<%
CurPage="调试任务"
strPage="atask"
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=4 width="<%=web_info(8)%>"  align="center">
		<Tr><Td class=ctd height=300>
			Welcome to 调试任务模块
		</Td></Tr>
	</Table>
<%
End Sub
%>