<!--#include file="include/conn.asp"-->
<%
CurPage="��������"
strPage="ftask"
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
			Welcome to ��������ģ��
		</Td></Tr>
	</Table>
<%
End Sub
%>