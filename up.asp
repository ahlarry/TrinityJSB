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
	Dim mystr, mystr1, ikhxs, rwlr_change, tmpRs
	ikhxs=1
	strSql="select * from [ftask] where datediff('m',jssj,'"&now()&"')<4 and (rwlx='��������' or rwlx='��������' or rwlx='�����������')"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Do while not Rs.eof
				if IsNull(Rs("ed")) Then Rs("ed")=0
				Rs.update
		Rs.movenext
		Loop
	Rs.close
end function
%>
