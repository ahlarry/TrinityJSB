<!--#include file="include/function.asp"-->
<!--#include file="inc/mtask_dbinf.asp"-->
<%
	Call ChkPageAble(4)
	pagelink=lnk_mtask
	session("pagelink")=pagelink
	web_curtitle="  → 删除任务书"
	xujian_ims.web_title="设计任务" & web_curtitle
	'xujian_ims.jsfiles_inc("js/mtask.js")
	xujian_ims.web_head()
	call lsh_search()
	response.write(hline(10, "100%", ""))

	if request.form("id") <> "" then
		call mtask_db_delete()	'删除库中内容操作
	else
		call main()
	end if

	response.write(hline(10, "100%", ""))
	xujian_ims.web_foot()

function main()
	dim s_lsh
	s_lsh=""
	if trim(request("s_lsh"))<>"" then s_lsh=trim(request("s_lsh"))
	if s_lsh="" then response.write(prompt("请输入要删除任务书的流水号!")) : exit function

	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	set rs=xujian_ims.exec(sql,1)
	if rs.eof or rs.bof then
		response.write(prompt("流水号为 <b>" & s_lsh & "</b> 的任务书不存在!"))
	else
		call mtask_delete(rs)
	end if
	rs.close
end function

function lsh_search()
%>
	<table border=0 cellpadding=2 cellspacing=0 width="100%">
		<form action=<%=request.servervariables("script_name")%> method=get>
		<tr>
			<td>&nbsp;&nbsp;
				输入删除任务书的流水号:<input type=text name=s_lsh size=8 value=<%=request("s_lsh")%>>
				<input type=submit value="查找">
			</td>
		</tr>
		<tr><td class=td_frame height=1></td></tr>
		</form>
	</table>
<%
end function

function mtask_delete(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>删除流水号 <font style=color:#0000FF><%=rs("lsh")%></font>    的任务书</font>
	<table class=table_xblue cellspacing=0 cellpadding=3 width="98%">
	<form id=mtask_add name=mtask_add action=<%=request.servervariables("script_name")%> method=post onSubmit='return confirm("任务书删除后将不能回复!\n您确信删除流水号 <%=rs("lsh")%> 的任务书吗?");'>

	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>■合同信息</b></td>
	</tr>

	<tr>
		<td class=td_rblue width="20%">订单号</td>
		<td class=td_lblue width="30%"><%=rs("ddh")%></td>
		<td class=td_rblue width="20%">流水号</td>
		<td class=td_lblue width="*"><%=rs("lsh")%></td>
	</tr>

	<tr>
		<td class=td_rblue>客户名称</td>
		<td class=td_lblue><%=rs("dwmc")%></td>
		<td class=td_rblue>断面名称</td>
		<td class=td_lblue><%=rs("dmmc")%></td>
	</tr>

	<tr>
		<td class=td_rblue>模号</td>
		<td class=td_lblue><%=rs("mh")%></td>
		<td class=td_rblue>设备厂家</td>
		<td class=td_lblue><%=rs("sbcj")%></td>
	</tr>

	<tr>
		<td class=td_rblue>挤出机型号</td>
		<td class=td_lblue><%=rs("jcjxh")%></td>
		<td class=td_rblue>水接头数量</td>
		<td class=td_lblue><%=rs("sjtsl")%></td>
	</tr>

	<tr>
		<td class=td_rblue>气接头数量</td>
		<td class=td_lblue><%=rs("qjtsl")%></td>
		<td class=td_rblue>配加热板</td>
		<td class=td_lblue><%if rs("pjrb") then%>是<%else%>否<%end if%></td>
	</tr>

	<tr>
		<td class=td_rblue>加热板信息</td>
		<td class=td_lblue>相数:<%=rs("jrbxs")%>	 材质:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
		<td class=td_rblue>模具材料</td>
		<td class=td_lblue><%=rs("mjcl")%></td>
	</tr>

	<tr>
		<td class=td_rblue>腔数</td>
		<td class=td_lblue><%=rs("qs")%>腔</td>
		<td class=td_rblue>牵引速度</td>
		<td class=td_lblue><%=rs("qysd")%>米/分(m/min)</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>■模具信息</b></td>
	</tr>

	<tr>
		<td class=td_rblue>任务内容</td>
		<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
		<td class=td_rblue>模头结构</td>
		<td class=td_lblue><%=rs("mtjg")%></td>
	</tr>

	<tr>
		<td class=td_rblue>定型结构</td>
		<td class=td_lblue><%=rs("dxjg")%>&nbsp;</td>
		<td class=td_rblue>水箱结构</td>
		<td class=td_lblue><%=rs("sxjg")%>&nbsp;</td>
	</tr>

	<tr>
		<td class=td_rblue>模头连接尺寸</td>
		<td class=td_lblue><%=rs("mtljcc")%>&nbsp;</td>
		<td class=td_rblue>热电偶规格</td>
		<td class=td_lblue><%=rs("rdogg")%>&nbsp;</td>
	</tr>


	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>■工艺信息</b></td>
	</tr>

	<tr>
		<td class=td_rblue>定型连接件出图</td>
		<td class=td_lblue><%if rs("dxljjct") then%>是<%else%>否<%end if%></td>
		<td class=td_rblue>定型切割</td>
		<td class=td_lblue><%=rs("dxqg")%></td>
	</tr>

	<tr>
		<td class=td_rblue>整体分流锥</td>
		<td class=td_lblue><%if rs("ztflz") then%>是<%else%>否<%end if%></td>
		<td class=td_rblue>整体型芯</td>
		<td class=td_lblue><%if rs("ztxx") then%>是<%else%>否><%end if%></td>
	</tr>

	<tr>
		<td class=td_rblue>整体定型块</td>
		<td class=td_lblue><%if rs("ztdxk") then%>是<%else%>否<%end if%></td>
		<td class=td_rblue>&nbsp;</td>
		<td class=td_lblue>&nbsp;</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>■其他信息</b></td>
	</tr>

	<tr>
		<td class=td_rblue >备注</td>
		<td class=td_lblue colspan=3><%=xujian_ims.htmltocode(rs("bz"))%></td>
	</tr>

	<tr>
		<td class=td_rblue>计划结束时间</td>
		<td class=td_lblue><%=rs("jhjssj")%></td>
		<td class=td_rblue>&nbsp;</td>
		<td class=td_lblue>&nbsp;</td>
	</tr>

	<tr>
		<td class=td_rblue>组长</td>
		<td class=td_lblue><%=rs("zz")%></td>
		<td class=td_rblue>技术代表</td>
		<td class=td_lblue><%=rs("jsdb")%></td>
	</tr>

	<tr><td class=td_cblue colspan=4><input type=submit value=" · 删除 · "></td></tr>
	<input type="hidden" name=id value=<%=rs("id")%>>
	<input type="hidden" name=s_lsh value=<%=rs("lsh")%>>
	</form>
	</table>
<%
end function		'mtask_delete()

function mtask_db_delete()
	dim iid, strtmplsh
	iid=request.form("id")
	strtmplsh=request("s_lsh")
	sql="delete from mtask_info where id=" & iid
	call xujian_ims.exec(sql, 0)
	sql="delete from mtask_flow where lsh='"&strtmplsh&"'"
	call xujian_ims.exec(sql, 0)
	err_title="任务书删除成功!"
	err_inf="流水号 <b>" & strtmplsh & "</b> 任务书删除成功!|||点击<a href=mtask_delete.asp>删除任务书</a>继续删除任务书!"
	gotoprompt(1)
	response.end
end function
%>
