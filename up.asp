<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="��������"					'ҳ�������λ��( ��������� �� ���������)
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
	slsh="12168,12172" : oldxs=1 : newxs=1.4
	slsh=split(slsh,",")
	for n=0 to ubound(slsh)
		oldxs=1
		strSql="select * from [mtask] where lsh='"&slsh(n)&"'"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		If Not Rs.eof Then
			oldxs=rs("fzxs")
			rs("fzxs")=newxs
		End If
		Rs.update
		Rs.close

		strSql="select * from [mantime] where Instr(rwlr,'����')=0 and lsh='"&slsh(n)&"'"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
			Do while not Rs.eof
				Response.write(Rs("lsh")&"---"&Rs("zrr")&"---"&Rs("fz")&"�� ")
				Rs("fz")=Round(Rs("fz")*newxs/oldxs,1)
				Rs.update
				Response.write(Rs("fz")&"<br>")
				Rs.movenext
			Loop
		Rs.close
	next
end function
%>
