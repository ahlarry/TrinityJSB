<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="��������"					'ҳ�������λ��( ��������� �� ���������)
strPage=""
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd height=300>
      <%Call Updata() %>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function Updata()
	Dim ijgsj, isj
	strSql="select * from [mtask] where datediff('m', sjjssj, '"&now()&"') < 3 or datediff('m', rwxdsj, '"&now()&"') < 4"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Do while not Rs.eof
			Rs("mtrw")=rs("rwlr")
			Rs("dxrw")=rs("rwlr")
			if Rs("mjxx")="����" Then Rs("mtrw")=""
			if Rs("mjxx")="ģͷ" Then Rs("dxrw")=""
			Rs.update
		Rs.movenext
		Loop
	Rs.close
end function
%>
