<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="升级数据"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage=""
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, strlsh
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
      <%
     	 Call UpFz()
      %>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function UpFz()
	Dim n, slsh, oldxs, newxs,TmpRs,TmpSql
	slsh="" : oldxs=0 : newxs=1
'	slsh=split(slsh,",")
'	for n=0 to ubound(slsh)
'		oldxs=1
'		strSql="select * from [mtask] where lsh='"&slsh(n)&"'"
'		Call xjweb.exec("",-1)
'		Rs.open strSql,Conn,1,3
'		If Not Rs.eof Then
'			oldxs=rs("fzxs")
'			rs("fzxs")=newxs
'		End If
'		Rs.update
'		Rs.close

'		strSql="select * from [mantime] where (Instr(rwlr,'结构')>0 or Instr(rwlr,'设计')>0 or Instr(rwlr,'BOM')>0) and lsh='"&slsh(n)&"'"
'		Call xjweb.exec("",-1)
'		Rs.open strSql,Conn,1,3
'			Do while not Rs.eof
'				Rs("fz")=Round(Rs("fz")*newxs/oldxs,1)
'				Rs.update
'				Rs.movenext
'			Loop
'		Rs.close
'	next

		TmpSql="select * from [mantime] where rwlr='模具复审确认'"
		Set TmpRs=Server.CreateObject("adodb.recordset")
		TmpRs.open TmpSql,Conn,1,3
		Do while not TmpRs.eof

'			TmpSql="select * from [mtask] where lsh='"&Rs("lsh")&"'"
'			Set TmpRs=Server.CreateObject("adodb.recordset")
'			TmpRs.open TmpSql,conn,1,3
'			If Not TmpRs.eof Then
'				TmpRs("fz")=0
'			End If
'			Response.Write(newxs&"、"&Rs("lsh")&"<br>")
'			TmpRs.update
'			TmpRs.close
			oldxs=0
			StrSql="select * from [mantime] where lsh='"&TmpRs("lsh")&"' and rwlr='模头设计确认'"
			Set Rs=Server.CreateObject("adodb.recordset")
			Rs.open StrSql,conn,1,3
			If Not Rs.eof Then
				oldxs=Rs("fz")
			End If
			Rs.close
			StrSql="select * from [mantime] where lsh='"&TmpRs("lsh")&"' and rwlr='定型设计确认'"
			Set Rs=Server.CreateObject("adodb.recordset")
			Rs.open StrSql,conn,1,3
			If Not Rs.eof Then
				oldxs=Rs("fz")+oldxs
			End If
			Rs.close
			oldxs=Round(oldxs/2,1)
		n=TmpRs("fz")
		if TmpRs("fz")<>0 and TmpRs("fz")<>oldxs Then TmpRs("fz")=oldxs
		TmpRs.update
		Response.Write(newxs&"、"&TmpRs("lsh")&" ："&n&"---"&oldxs&"---"&TmpRs("fz")&"<br>")
		TmpRs.movenext
		newxs=newxs+1
		Loop
		TmpRs.close
end function
%>
