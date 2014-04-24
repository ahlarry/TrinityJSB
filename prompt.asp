<!--#include file="include/conn.asp"-->
<%
CurPage="系统提示"
Dim iCode, strTitle, strContents, strPreUrl, strNewURl
iCode=Session("InfoCode")
strTitle=Session("InfoTitle")
strContents=Session("InfoContens")
strPreUrl=Session("InfoPreUrl")
strNewURl=Session("InfoNewUrl")

Session("InfoCode")=""
Session("InfoTitle")=""
Session("InfoContens")=""
Session("InfoPreUrl")=""
Session("InfoNewUrl")=""

Call FileInc(0, "str")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellpadding=0 cellspacing=0 width="<%=web_info(8)%>" align="center">
		<Tr><Td class=ctd height=300>
			<Table class=xtable cellpadding=5 cellspacing=0 width="300"  align="center">
				<Tr><Td class=th align=left><font style=color:yellow>系统温馨提示:</font> <%=strTitle %></Td></Tr>
				<Tr><Td class=ltd><%=strContents %></Td></Tr>
			</Table>
		<Td></Tr>
	<Table>
<%
End Sub
%>