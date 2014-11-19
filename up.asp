<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="升级数据"					'页面的名称位置( 任务书管理 → 添加任务书)
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
	strSql="select * from [ftask] where datediff('m',jssj,'"&now()&"')<4 and (rwlx='零星修理' or rwlx='零星任务' or rwlx='技术代表设计')"
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
