<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
Call ChkPageAble(0)
CurPage="������� �� �鿴������"
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
	If s_lsh="" Then Call TbTopic("������鿴���������ˮ��!") : Exit Function
	strSql="select * from [mtask] where lsh = '"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("��ˮ�� ��" & s_lsh & "�� �����鲻����! ������������ˮ��!", "")
	Else
		Select Case action
			Case "max"
				Call mtask_muchinfo(Rs)
				Response.Write(xjline(5, "100%", ""))
				If Session("userdepart")="������" Then
					Call mtask_technicsinfo(rs)
					Response.Write(XjLine(5, "100%", ""))
				End if
				Call mtask_alluserinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Call atask_alluserinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Response.Write("<a href=?s_lsh="&rs("lsh")&">��Ҫ��Ϣ</a> &nbsp;")
				Response.Write("<a href=?action=min&s_lsh="&rs("lsh")&">������Ϣ</a> &nbsp;")
			Case "min"
				Call mtask_fewinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Call mtask_userinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Call atask_userinfo(rs)
				Response.Write(XjLine(5, "100%", ""))
				Response.Write("<a href=?action=max&s_lsh="&rs("lsh")&">������Ϣ</a> &nbsp;")
				Response.Write("<a href=?s_lsh="&rs("lsh")&">��Ҫ��Ϣ</a> &nbsp;")
			Case else
				If not Session("userdepart")="������" Then
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
					Response.Write("<a href=?action=max&s_lsh="&rs("lsh")&">������Ϣ</a> &nbsp;")
					Response.Write("<a href=?action=min&s_lsh="&rs("lsh")&">������Ϣ</a> &nbsp;")
				End if
		End Select
	End If
	Rs.Close
End Function
%>