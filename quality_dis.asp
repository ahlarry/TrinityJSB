<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="问题分析 → 查看外部质量信息"
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
			<%Call qualityDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function qualityDisplay()
	Dim s_hth, iid, strAlt
	s_hth="" : iid=0 : strAlt=""
	s_hth=Trim(Request("s_lsh"))
	If IsNumeric(Trim(Request("id"))) Then iid=CInt(Trim(Request("id")))
	If s_hth="" Then Call TbTopic("请输入查看外部质量信息的合同号!") : Exit Function
	strSql="select lxr, lxdh, gzlh,jssj from [quality] where hth='"&s_hth&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Not(Rs.Eof Or Rs.Bof) Then
		strAlt="联系人:" & Rs("lxr") & "<br>联系电话:" & Rs("lxdh") & "<br>工作令号" & Rs("gzlh") & "<br>接收时间:" & Rs("jssj")
	End If
	Rs.Close

	strSql="select * from [quality] where hth='"&s_hth&"'"
	If iid<>0 Then strSql=strSql & " and id="&iid&""
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("合同号 【" & s_lsh & "】 没有任何问题! ","quality_list.asp")
	Else
		Call quality_Display(Rs, strAlt)
	End If
	Rs.close
End Function

Function Task_Info(Rs)
%>
	<%Call TbTopic("合同号 "&Rs("hth")&" 质量信息")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
		<tr>
			<td class=th width="20%">联系人</td>
			<td class=th width="*">联系电话</td>
			<td class=th width="20%">工作令号</td>
			<td class=th width="20%">接收时间</td>
		</tr>

		<tr>
			<td class=ctd><%=Rs("lxr")%></a></td>
			<td class=ctd><%=Rs("lxdh")%></td>
			<td class=ctd><%=Rs("gzlh")%></td>
			<td class=ctd><%=Rs("jssj")%></td>
		</tr>
	</table>
<%
End Function

Function quality_Display(Rs, strAlt)
	Call TbTopic("合同号 " &Rs("hth")&" 外部质量信息")
%>
	<%Do While Not Rs.Eof%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
		<Tr>
			<Td class=th width=80>合同号</Td>
			<Td class=ctd width=80 <%If strAlt<>"" Then%>alt="<%=strAlt%>"<%End If%>><%=Rs("hth")%></Td>
			<Td class=th width=80>客户名称</Td>
			<Td class=ctd width=*><%=Rs("khmc")%></Td>
			<Td class=th width=80>责任人</Td>
			<Td class=ctd width=80><%=Rs("zrr")%></Td>
			<Td class=th width=80>当前状态</Td>
			<Td class=ctd width=80><%=Rs("wczk")%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>主要问题</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("zywt"))%></Td>
		</Tr>

		<% If Rs("yjcs") <> "" Then %>
		<Tr>
			<Td class=th width=80>应急措施</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("yjcs"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("yyfx") <> "" Then %>
		<Tr>
			<Td class=th width=80>原因分析</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("yyfx"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("jzcs") <> "" Then %>
		<Tr>
			<Td class=th width=80>纠正措施</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("jzcs"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("lsqk") <> "" Then %>
		<Tr>
			<Td class=th width=80>落实情况</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("lsqk"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("yzjl") <> "" Then %>
		<Tr>
			<Td class=th width=80>验证结论</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("yzjl"))%></Td>
		</Tr>
		<% End If %>

		<%If ChkAble(11) Then%>
		<Tr>
			<Td colspan=4 class=ctd><a href="quality_change.asp?id=<%=Rs("id")%>">更改</a></Td>
			<Td colspan=4 class=ctd><a href="quality_indb.asp?action=delete&id=<%=Rs("id")%>" onclick="return confirm('确认删除吗?');">删除</a></Td>
		</Tr>
		<%End If%>
		</Table>
	<%
			Response.write(XjLine(10, "100%", ""))
			Rs.MoveNext
		Loop
End Function
%>