<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'11:45 2007-1-8-星期一
Call ChkPageAble(3)
CurPage="设计任务 → 删除任务书"
strPage="mtask"
'Call FileInc(0, "js/mtask.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
	Dim iid
	iid=Request.Form("id")
	If iid<>"" Then
		Call mtask_db_delete()
	Else
	%>
		<Table class=xtable cellspacing=0 cellpadding=2 width="<%=web_info(8)%>">
			<Tr><Td class=ctd>
				<%Call SearchLsh()%>
			</td></tr>
			<Tr><Td class=ctd height=300>
				<%Call mtaskDelete()%>
				<%Response.Write(XjLine(10,"100%",""))%>
			</Td></Tr>
		</Table>
	<%
	End If
End Sub

Function mtaskDelete()
	Dim s_lsh
	s_lsh=""
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("请输入要删除任务书的流水号!") : Exit Function

	strSql="select * from [mtask] where [lsh]='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("流水号为 【" & s_lsh & "】 的任务书不存在!","mtask_delete.asp")
	Else
		If Not IsNull(rs("sjjssj")) Then
			Call JsAlert("流水号为 【" & s_lsh & "】 的任务书已经完成,不能删除!","mtask_delete.asp")
		Else
			Call mtask_Delete(Rs)
		End If
	End If
	Rs.Close
End Function

Function mtask_delete(rs)
%>
	<%Call TbTopic("删除流水号 <font style=color:#0000FF>" &rs("lsh")&"</font> 的任务书")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%">
	<form action=<%=request.servervariables("script_name")%> method=post onSubmit='return confirm("任务书删除后将不能回复!\n您确信删除流水号 【<%=rs("lsh")%>】 的任务书吗?");'>

	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>■合同信息</b></td>
	</tr>

	<tr>
		<td class=rtd width="20%">订单号</td>
		<td class=ltd width="30%"><%=rs("ddh")%></td>
		<td class=rtd width="20%">流水号</td>
		<td class=ltd width="*"><%=rs("lsh")%></td>
	</tr>

	<tr>
		<td class=rtd>客户名称</td>
		<td class=ltd><%=rs("dwmc")%></td>
		<td class=rtd>断面名称</td>
		<td class=ltd><%=rs("dmmc")%></td>
	</tr>

	<tr>
		<td class=rtd>模号</td>
		<td class=ltd><%=rs("mh")%></td>
		<td class=rtd>设备厂家</td>
		<td class=ltd><%=rs("sbcj")%></td>
	</tr>

	<tr>
		<td class=rtd>挤出机型号</td>
		<td class=ltd><%=rs("jcjxh")%></td>
		<td class=rtd>水接头数量</td>
		<td class=ltd><%=rs("sjtsl")%></td>
	</tr>

	<tr>
		<td class=rtd>气接头数量</td>
		<td class=ltd><%=rs("qjtsl")%></td>
		<td class=rtd>配加热板</td>
		<td class=ltd><%if rs("pjrb") then%>是<%else%>否<%end if%></td>
	</tr>

	<tr>
		<td class=rtd>加热板信息</td>
		<td class=ltd>相数:<%=rs("jrbxs")%>	 材质:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
		<td class=rtd>模具材料</td>
		<td class=ltd><%=rs("mjcl")%></td>
	</tr>

	<tr>
		<td class=rtd>腔数</td>
		<td class=ltd><%=rs("qs")%>腔</td>
		<td class=rtd>牵引速度</td>
		<td class=ltd><%=rs("qysd")%>米/分(m/min)</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>■模具信息</b></td>
	</tr>

	<tr>
		<td class=rtd>任务内容</td>
		<td class=ltd><%=rs("mjxx") & rs("rwlr")%></td>
		<td class=rtd>模头结构</td>
		<td class=ltd><%=rs("mtjg")%></td>
	</tr>

	<tr>
		<td class=rtd>定型结构</td>
		<td class=ltd><%=rs("dxjg")%>&nbsp;</td>
		<td class=rtd>水箱结构</td>
		<td class=ltd><%=rs("sxjg")%>&nbsp;</td>
	</tr>

	<tr>
		<td class=rtd>模头连接尺寸</td>
		<td class=ltd><%=rs("mtljcc")%>&nbsp;</td>
		<td class=rtd>热电偶规格</td>
		<td class=ltd><%=rs("rdogg")%>&nbsp;</td>
	</tr>


	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>■工艺信息</b></td>
	</tr>

	<tr>
		<td class=rtd>定型连接件出图</td>
		<td class=ltd><%if rs("dxljjct") then%>是<%else%>否<%end if%></td>
		<td class=rtd>定型切割</td>
		<td class=ltd><%=rs("dxqg")%></td>
	</tr>

	<tr>
		<td class=rtd>整体分流锥</td>
		<td class=ltd><%if rs("ztflz") then%>是<%else%>否<%end if%></td>
		<td class=rtd>整体型芯</td>
		<td class=ltd><%if rs("ztxx") then%>是<%else%>否><%end if%></td>
	</tr>

	<tr>
		<td class=rtd>整体定型块</td>
		<td class=ltd><%if rs("ztdxk") then%>是<%else%>否<%end if%></td>
		<td class=rtd>&nbsp;</td>
		<td class=ltd>&nbsp;</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>■其他信息</b></td>
	</tr>

	<tr>
		<td class=rtd >评审记录</td>
		<td class=ltd colspan=3><%=xjweb.HtmlToCode(Rs("psjl"))%></td>
	</tr>

	<tr>
		<td class=rtd >备注</td>
		<td class=ltd colspan=3><%=xjweb.HtmlToCode(Rs("bz"))%></td>
	</tr>

	<tr>
		<td class=rtd>计划结束时间</td>
		<td class=ltd><%=rs("jhjssj")%></td>
		<td class=rtd>&nbsp;</td>
		<td class=ltd>&nbsp;</td>
	</tr>

	<tr>
    <td class=rtd>组长</td>
    <td class=ltd>
    <%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(结构)、"&rs("sjzz")&"(设计)")%>
    </td>
		<td class=rtd>技术代表</td>
		<td class=ltd><%=rs("jsdb")%></td>
	</tr>

	<tr><td class=ctd colspan=4><input type=submit value=" · 删除 · "></td></tr>
	<input type="hidden" name=id value=<%=rs("id")%>>
	<input type="hidden" name=s_lsh value=<%=rs("lsh")%>>
	</form>
	</table>
<%
End Function		'mtask_delete()

Function mtask_db_delete()
	Dim iid, strlsh
	iid=Request.Form("id")
	strlsh=Request.Form("s_lsh")
	strSql="delete from [mtask] where [id]=" & iid
	Call xjweb.Exec(strSql, 0)
	Call JsAlert("流水号 【 " & strlsh & " 】 任务书删除成功!", "mtask_delete.asp")
End Function
%>
