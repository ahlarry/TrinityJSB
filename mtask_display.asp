<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
Call ChkPageAble(0)
CurPage="设计任务 → 查看任务书"
strPage="mtask"
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=2 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchLsh()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call mtaskDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function mtaskDisplay()
	Dim s_lsh, action
	s_lsh=Trim(Request("s_lsh"))
	action=Trim(Request("action"))
	If s_lsh="" Then Call TbTopic("请输入查看任务书的流水号!") : Exit Function
	strSql="select * from [mtask] where lsh = '"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书不存在! 请重新输入流水号!", "")
	Else
		Select Case action
			Case "max"
				Call mtask_muchinfo(Rs)
				Response.Write(xjline(5, "100%", ""))
				If Session("userdepart")="技术部" Then
					Call mtask_technicsinfo(rs)
					Response.Write(XjLine(5, "100%", ""))
				End if
				Call mtask_alluserinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Call atask_alluserinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Response.Write("<a href=?s_lsh="&rs("lsh")&">主要信息</a> &nbsp;")
				Response.Write("<a href=?action=min&s_lsh="&rs("lsh")&">部分信息</a> &nbsp;")
			Case "min"
				Call mtask_fewinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Call mtask_userinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Call atask_userinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Response.Write("<a href=?action=max&s_lsh="&rs("lsh")&">所有信息</a> &nbsp;")
				Response.Write("<a href=?s_lsh="&rs("lsh")&">主要信息</a> &nbsp;")
			Case else
				If not Session("userdepart")="技术部" Then
					call mtask_technicsinfo(rs)
					Response.Write(XjLine(5, "100%", ""))
				else
					Call mtask_muchinfo(rs)
					Response.Write(XjLine(5, "100%", ""))
					call mtask_technicsinfo(rs)
					Response.Write(XjLine(5, "100%", ""))
					Call mtask_userinfo(rs)
					Response.Write(XjLine(5, "100%", ""))
					Call atask_userinfo(rs)
					Response.Write(XjLine(5, "100%", ""))
					Response.Write("<a href=?action=max&s_lsh="&rs("lsh")&">所有信息</a> &nbsp;")
					Response.Write("<a href=?action=min&s_lsh="&rs("lsh")&">部分信息</a> &nbsp;")
				End if
		End Select
	End If
	Rs.Close
End Function
%>