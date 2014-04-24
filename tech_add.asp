<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(7)
CurPage="问题分析 → 添加问题分析"
strPage="tech"
Call FileInc(0, "js/tech.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%Call techAdd()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

function techAdd()
%>
	<%Call TbTopic("添加技术问题分析")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<form id=frm_techadd name=frm_techadd action=tech_indb.asp?action=add method=post onSubmit='return checkinf();'>

	<tr>
		<th class=th height=25>项目名称</td>
		<th class=th>项目内容</td>
	</tr>

	<tr>
		<td class=rtd>流水号</td>
		<td class=ltd><input type=text name="lsh" size=15></td>
	</tr>
	<tr>
		<td class=rtd>板块名称</td>
		<td class=ltd><input type=text name="bkmc" size=15></td>
	</tr>

	<tr>
		<td class=rtd width="20%">处理意见</td>
		<td class=ltd>
			<select name="clyj">
				<option value="报废">报废</option>
				<option value="返修">返修</option>
				<option value="返修">留用</option>
			</select>
		</td>
	</tr>

	<tr>
		<td class=rtd>责任人</td>
		<td class=ltd><input type=text name="zrr" size=15></td>
	</tr>

	<tr>
		<td class=rtd valign=top>问题现象描述</td>
		<td class=ltd><textarea name="xxms" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd valign=top>产生原因分析</td>
		<td class=ltd><textarea name="yyfx" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd valign=top>纠正预防措施</td>
		<td class=ltd><textarea name="yfcs" cols="75" rows="7"></textarea></td>
	</tr>

	<tr><td class=ctd colspan=2><input type=submit value=" · 确 定 · "></td></tr>
	</form>
	</table>
<%
end function		'tech_add()
%>