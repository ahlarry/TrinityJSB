<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(6)
CurPage="模具调试 → 更改调试信息"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="mtest"
Call FileInc(0, "js/mtest.js")
xjweb.header()
Call TopTable()

Dim iid, strlsh, ics, bps
iid=clng(Request("id"))
strlsh=Request("s_lsh")
ics=clng(Request("cs"))
bps=Request("ps")

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
			<%Call mtestChange()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function mtestChange()
	If strlsh="" Then Call JsAlert("请从正规入口!","mtest_list.asp") : Exit Function
	strSql="select * from [ts_tsxx] where id=" & iid
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("出错啦!请从正规入口!","mtest_list.asp") : Exit Function
	Else
		Response.Write(xjLine(10, "100%", ""))
		If bps Then
			Call mtestps_change(Rs)
		Else
			Call mtest_change(Rs)
		End If
	End If
	Rs.close
End Function


Function mtest_change(Rs)
%>
	<%Call TbTopic("更改流水号 "&strlsh&" 模具第 " & ics &" 次调试信息")%>
	<table class=xtable cellspacing=0 cellpadding=3 align="center">
	<form id=frm_mtestadd name=frm_mtestadd action=mtest_indb.asp?action=change method=post onSubmit='return tscheckinf();'>
	<tr>
		<td class=rtd>调试原因</td>
		<td class=ltd colspan=6><textarea name="tsyy" cols="90" rows="7"><%=Rs("tsyy")%></textarea></td>
	</tr>
	<tr>
		<td class=rtd>调试内容</td>
		<td class=ltd colspan=6><textarea name="tslr" cols="90" rows="7"><%=Rs("tslr")%></textarea></td>
	</tr>
	<tr>
		<td class=ctd >挤出机</td>
		<td class=ctd >螺筒1</td>
		<td class=ctd >螺筒2</td>
		<td class=ctd >螺筒3</td>
		<td class=ctd >螺筒4</td>
		<td class=ctd >螺筒5</td>
		<td class=ctd >螺筒6</td>
	</tr>	
	<tr>
		<td class=ctd ><input name="jcj" size=8 value=<%=Rs("jcj")%>></td>
		<td class=ctd ><input name="lt1" size=8 value=<%=Rs("lt1")%>></td>
		<td class=ctd ><input name="lt2" size=8 value=<%=Rs("lt2")%>></td>
		<td class=ctd ><input name="lt3" size=8 value=<%=Rs("lt3")%>></td>
		<td class=ctd ><input name="lt4" size=8 value=<%=Rs("lt4")%>></td>
		<td class=ctd ><input name="lt5" size=8 value=<%=Rs("lt5")%>></td>
		<td class=ctd ><input name="lt6" size=8 value=<%=Rs("lt6")%>></td>
	</tr>	
	<tr>
		<td class=ctd >过渡体</td>
		<td class=ctd >模头温度</td>
		<td class=ctd >螺杆速度</td>
		<td class=ctd >加料速度</td>
		<td class=ctd >扭矩</td>		
		<td class=ctd >融温</td>
		<td class=ctd >融压</td>
	</tr>	
	<tr>
		<td class=ctd ><input name="gdt" size=8 value=<%=Rs("gdt")%>></td>
		<td class=ctd ><input name="mtwd" size=8 value=<%=Rs("mtwd")%>></td>
		<td class=ctd ><input name="lgsd" size=8 value=<%=Rs("lgsd")%>></td>
		<td class=ctd ><input name="jlsd" size=8 value=<%=Rs("jlsd")%>></td>
		<td class=ctd ><input name="niuj" size=8 value=<%=Rs("niuj")%>></td>		
		<td class=ctd ><input name="rongw" size=8 value=<%=Rs("rongw")%>></td>
		<td class=ctd ><input name="rongy" size=8 value=<%=Rs("rongy")%>></td>
	</tr>		
	<tr><td class=ctd colspan=7><input type=submit value=" · 更 改 · "></td></tr>
	<input type="hidden" name="id" value=<%=iid%>>
	<input type="hidden" name="lsh" value=<%=Rs("lsh")%>>
	</form>
	</table>
<%
End Function		'mtest_change()

Function mtestps_change(Rs)
%>
	<%Call TbTopic("添加流水号 " & strlsh & " 模具第 " & ics & " 次评审记录") %>
	<table class=xtable cellspacing=0 cellpadding=3 width="98%">
	<form id=frm_mtestpsadd name=frm_mtestpsadd action=mtest_indb.asp?action=change method=post onSubmit='return tspscheckinf();'>

	<tr>
		<th class=rtd height=25 width="20%">项目名称</td>
		<th class=ctd width="*">项目内容</td>
	</tr>
	<tr>
		<td class=rtd>评审内容</td>
		<td class=ltd><textarea name="tslr" cols="95" rows="7"><%=Rs("tslr")%></textarea></td>
	</tr>

	<tr>
		<td class=rtd>评审人</td>
		<td class=ltd><textarea name="tsyy" cols="95" rows="3"><%=Rs("tsyy")%></textarea></td>
	</tr>
	<tr><td class=ctd colspan=2><input type=submit value=" · 更 改 · "></td></tr>
	<input type="hidden" name="id" value=<%=iid%>>
	<input type="hidden" name="lsh" value=<%=Rs("lsh")%>>
	</form>

	</table>
<%
End Function
%>