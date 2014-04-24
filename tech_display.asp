<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="问题分析 → 查看问题分析"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="tech"
'Call FileInc(0, "js/mtest.js")
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
			<%Call techDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function techDisplay()
	Dim s_lsh, iid, strAlt
	s_lsh="" : iid=0 : strAlt=""
	s_lsh=Trim(Request("s_lsh"))
	If IsNumeric(Trim(Request("id"))) Then iid=CInt(Trim(Request("id")))
	If s_lsh="" Then Call TbTopic("请输入查看问题分析的模具流水号!") : Exit Function
	strSql="select lsh, dmmc, mh,dwmc from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Not(Rs.Eof Or Rs.Bof) Then
		strAlt="流水号:" & Rs("lsh") & "<br>单位名称:" & Rs("dwmc") & "<br>断面名称" & Rs("dmmc") & "<br>模号:" & Rs("mh")
	End If
	Rs.Close

	strSql="select * from [tecq_question] where tecq_lsh='"&s_lsh&"'"
	If iid<>0 Then strSql=strSql & " and id="&iid&""
	'Response.Write strSql
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书没有任何问题! ","tech_list.asp")
	Else
		Call tech_Display(Rs, strAlt)
	End If
	Rs.close
End Function

Function Task_Info(Rs)
%>
	<%Call TbTopic("流水号 "&Rs("lsh")&" 模具信息")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%">
		<tr> 
			<td class=th width="20%">流水号</td>
			<td class=th width="*">断面名称</td>
			<td class=th width="20%">模号</td>
			<td class=th width="20%">单位名称</td>
		</tr>
								
		<tr> 
			<td class=ctd><a href="mtask_display.asp?s_lsh=<%=Rs("lsh")%>" alt="查看流水号 <b><%=Rs("lsh")%></b> 任务书"><%=Rs("lsh")%></a></td>
			<td class=ctd><%=Rs("dmmc")%></td>
			<td class=ctd><%=Rs("mh")%></td>
			<td class=ctd><%=Rs("dwmc")%></td>
		</tr>
	</table>	
<%
End Function

Function tech_Display(Rs, strAlt)
	Call TbTopic("流水号 " &Rs("tecq_lsh")&" 模具技术问题分析")
%>
	<%Do While Not Rs.Eof%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="95%">
		<Tr>
			<Td class=th width=80>流水号</Td>
			<Td class=ctd width=80 <%If strAlt<>"" Then%>alt="<%=strAlt%>"<%End If%>><a href="mtask_display.asp?s_lsh=<%=Rs("tecq_lsh")%>" alt="查看流水号 <b><%=Rs("tecq_lsh")%></b> 任务书"><%=Rs("tecq_lsh")%></a></Td>
			<Td class=th width=80>板块名称</Td>
			<Td class=ctd width=*><%=Rs("tecq_bkmc")%></Td>
			<Td class=th width=80>模号</Td>
			<Td class=ctd width=80><%=Rs("tecq_clyj")%></Td>
			<Td class=th width=80>责任人</Td>
			<Td class=ctd width=80><%=Rs("tecq_zrr")%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>问题现象描述</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("tecq_xxms"))%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>产生原因分析</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("tecq_yyfx"))%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>纠正预防措施</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("tecq_yfcs"))%></Td>
		</Tr>
		<%If ChkAble("1,7") Then%>
		<Tr>
			<Td colspan=4 class=ctd><a href="tech_change.asp?id=<%=Rs("id")%>">更改</a></Td>
			<Td colspan=4 class=ctd><a href="tech_indb.asp?action=delete&id=<%=Rs("id")%>" onclick="return confirm('确认删除吗?');">删除</a></Td>
		</Tr>
		<%End If%>
		</Table>
	<%	
			Response.write(XjLine(10, "100%", ""))
			Rs.MoveNext
		Loop
	%>
	
	
<%
End Function
%>