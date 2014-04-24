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
	Dim ijgsj, isj
	strSql="select * from [mtask] where isnull(sjjssj) and isnull(jhjgsj) order by jhjssj desc"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Do while not Rs.eof
			isj=INT(datediff("d", rs("jhkssj"), rs("jhjssj"))/2)
			ijgsj=dateadd("d",isj,rs("jhkssj"))
			Rs("jhjgsj")=ijgsj
			Rs.update
			Response.Write(Rs("lsh")&"-----"&Rs("jhjgsj")&"<br>")
		Rs.movenext
		Loop
	Rs.close
end function
%>
