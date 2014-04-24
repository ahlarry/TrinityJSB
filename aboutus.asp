<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="关于我们"
strPage=""	 'aboutus
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
			<font style="font-size:15pt;font-weight:bold;">我们的集体</font>
			<img src="<%=web_info(2)%>images/collectivity.JPG" onload="javascript:if(this.width>(<%=web_info(8)%>-10)) this.width=(<%=web_info(8)%>-10);" border="0"></img>
		</Td></Tr>
	</Table>
<%
End Sub
%>