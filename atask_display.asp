<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="�������� �� �鿴��������"					'ҳ�������λ��( ��������� �� ���������)
strPage="atask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchLsh()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call ataskDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function ataskDisplay()
	Dim s_lsh
	s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("������鿴�����������ˮ��!") : Exit Function
	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	Set rs=xjweb.Exec(strSql,1)
	if rs.eof or rs.bof then
		Call JsAlert("��ˮ�� ��" & s_lsh & "�� ���������񲻴���!","atask_display.asp")
	else
		call mtask_fewinfo(rs)
		response.write(xjLine(4, "100%", ""))
		call atask_userinfo(rs)
	end if
	rs.close
end function
%>