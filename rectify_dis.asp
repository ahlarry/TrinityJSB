<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="问题分析 → 查看纠正/预防措施表"
strPage="tech"
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
			<%Call RectifyDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function RectifyDisplay()
	Dim s_hth, iid
	s_hth="" : iid=0
	s_hth=Trim(Request("s_lsh"))
	If IsNumeric(Trim(Request("id"))) Then iid=CInt(Trim(Request("id")))
	If s_hth="" Then Call TbTopic("请输入查看纠正/预防措施表的编号!") : Exit Function

	strSql="select * from [Rectify] where bh='"&s_hth&"'"
	If iid<>0 Then strSql=strSql & " and id="&iid&""
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("编号 【" & s_lsh & "】 纠正/预防措施表不存在! ","Rectify_list.asp")
	Else
		Call Rectify_Display(Rs)
	End If
	Rs.close
End Function

Function Rectify_Display(Rs)
	Call TbTopic("编号 " &Rs("bh")&" 纠正/预防措施表")
%>
	<%Do While Not Rs.Eof%>
		<Table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
		<tr>
			<td class=th width="20%">责任部门</td>
			<Td class=ctd><%=Rs("zrbm")%></Td>
			<td class=th width="20%" colspan="2">编号</td>
			<Td class=ctd><%=Rs("bh")%></Td>
		</tr>

		<tr>
			<td class=th>发出信息部门</td>
			<Td class=ctd><%=Rs("xxbm")%></Td>
			<td class=th colspan="2">信息发出日期</td>
			<Td class=ctd><%=Rs("jssj")%></Td>
		</tr>

		<tr>
			<td class=th>不合格/潜在不合格内容</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("bhgnr"))%></td>
		</tr>

		<% If Rs("yfcsyq") <> "" Then %>
		<tr>
			<td class=th rowspan="2">纠正/预防措施要求</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("yfcsyq"))%></td>
		</tr>

		<tr>
			<td class=ctd colspan="2">期限:<%=Rs("qxsj")%></td>
			<td class=ctd colspan="2">是否评审：<% if Rs("ps")="V1" Then %>是<% else %>否<% End if %></td>
		</tr>
		<% End If %>
		<% If Rs("yyfx") <> "" Then %>
		<tr>
			<td class=th>原因分析</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("yyfx"))%></td>
		</tr>
		<% End If %>
		<% If Rs("jzcs") <> "" Then %>
		<tr>
			<td class=th>纠正措施</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("jzcs"))%></td>
		</tr>
		<% End If %>
		<% If Rs("lsqk") <> "" Then %>
		<tr>
			<td class=th>落实情况</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("lsqk"))%></td>
		</tr>
		<% End If %>
		<% If Rs("yzjl") <> "" Then %>
		<tr>
			<td class=th>验证结论</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("yzjl"))%></td>
		</tr>
		<% End If %>
		<%If ChkAble(11) Then%>
		<tr>
			<td class=ctd colspan="2" align="center"><a href="rectify_change.asp?id=<%=Rs("id")%>">更改</a></td>
			<td class=ctd colspan="3" align="center"><a href="rectify_indb.asp?action=delete&id=<%=Rs("id")%>" onclick="return confirm('确认删除吗?');">删除</a></td>
		</tr>
		<% End If %>
	</table>
	<%
			Response.write(XjLine(10, "100%", ""))
			Rs.MoveNext
		Loop
End Function
%>