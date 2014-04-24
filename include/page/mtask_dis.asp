<%
function mtask_display_few(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>挤出模具厂挤出模设计任务书</font>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4><b>■合同信息■</b></td>
		</tr>

		<tr>
			<td class=td_rblue width="15%">订单号</td>
			<td class=td_lblue width="35%"><%=rs("ddh")%></td>
			<td class=td_rblue width="15%">流水号</td>
			<td class=td_lblue width="*"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
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
			<td class=td_rblue>任务内容</td>
			<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
			<td class=td_rblue>计划结束时间</td>
			<td class=td_lblue><%=rs("jhjssj")%></td>
		</tr>
	</table>
<%
end function

function mtask_display_much(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>挤出模具厂挤出模设计任务书</font>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4><b>合同信息</b></td>
		</tr>

		<tr>
			<td class=td_rblue width="13%">订单号</td>
			<td class=td_lblue width="37%"><%=rs("ddh")%></td>
			<td class=td_rblue width="13%">流水号</td>
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
			<td class=td_rblue>模具材料</td>
			<td class=td_lblue><%=rs("mjcl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>设备厂家</td>
			<td class=td_lblue><%=rs("sbcj")%></td>
			<td class=td_rblue>水接头数量</td>
			<td class=td_lblue><%=rs("sjtsl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>挤出机型号</td>
			<td class=td_lblue><%=rs("jcjxh")%></td>
			<td class=td_rblue>气接头数量</td>
			<td class=td_lblue><%=rs("qjtsl")%></td>
			
		</tr>
		<tr>
			<td class=td_rblue>配加热板</td>
			<td class=td_lblue><%if rs("pjrb") then%>是<%else%>否<%end if%></td>
			<td class=td_rblue>腔数</td>
			<td class=td_lblue>
				<%if rs("qs")=1 then%>单腔<%end if%>
				<%if rs("qs")=2 then%>双腔<%end if%>
				<%if rs("qs")=3 then%>三腔<%end if%>
				<%if rs("qs")=4 then%>四腔<%end if%>
				<%if rs("qs")=5 then%>五腔<%end if%>
				<%if rs("qs")=6 then%>六腔<%end if%>
				<%if rs("qs")=7 then%>七腔<%end if%>
				<%if rs("qs")=8 then%>八腔<%end if%>
			</td>
		</tr>
		
		<tr>
			<td class=td_rblue>加热板信息</td>
			<td class=td_lblue>相数:<%=rs("jrbxs")%>	 材质:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
			<td class=td_rblue>牵引速度</td>
			<td class=td_lblue><%=rs("qysd")%> 米/分(m/min)</td>
		</tr>

		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>模具信息</b></td>
		</tr>

		<tr>
			<td class=td_rblue>模头结构</td>
			<td class=td_lblue><%=rs("mtjg")%>&nbsp;</td>
			<td class=td_rblue>任务内容</td>
			<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
		</tr>

		<tr>
			<td class=td_rblue>定型结构</td>
			<td class=td_lblue><%=rs("dxjg")%>&nbsp;</td>
			<td class=td_rblue>模头连接尺寸</td>
			<td class=td_lblue><%=rs("mtljcc")%></td>
		</tr>

		<tr>
			<td class=td_rblue>水箱结构</td>
			<td class=td_lblue><%=rs("sxjg")%>&nbsp;</td>
			<td class=td_rblue>热电偶规格</td>
			<td class=td_lblue><%=rs("rdogg")%></td>
		</tr>


		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>工艺信息</b></td>
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
			<td class=td_lblue height=25 colspan=4> <b>其他信息</b></td>
		</tr>

		<tr>
			<td class=td_rblue >评审记录</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("psjl"))%></td>
		</tr>

		<tr>
			<td class=td_rblue >备注</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("bz"))%></td>
		</tr>

		<tr>
			<td class=td_rblue>计划结束时间</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
			<td class=td_rblue>实际结束时间;</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
		</tr>

		<tr>
			<td class=td_rblue>组长</td>
			<td class=td_lblue><%=rs("zz")%></td>
			<td class=td_rblue>技术代表</td>
			<td class=td_lblue><%=rs("jsdb")%></td>
		</tr>
	</table>
<%
end function
 
function mtask_display_all(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>挤出模具厂挤出模设计任务书</font>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4><b>合同信息</b></td>
		</tr>

		<tr>
			<td class=td_rblue width="13%">订单号</td>
			<td class=td_lblue width="37%"><%=rs("ddh")%></td>
			<td class=td_rblue width="13%">流水号</td>
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
			<td class=td_rblue>模具材料</td>
			<td class=td_lblue><%=rs("mjcl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>设备厂家</td>
			<td class=td_lblue><%=rs("sbcj")%></td>
			<td class=td_rblue>水接头数量</td>
			<td class=td_lblue><%=rs("sjtsl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>挤出机型号</td>
			<td class=td_lblue><%=rs("jcjxh")%></td>
			<td class=td_rblue>气接头数量</td>
			<td class=td_lblue><%=rs("qjtsl")%></td>
			
		</tr>
		<tr>
			<td class=td_rblue>配加热板</td>
			<td class=td_lblue><%if rs("pjrb") then%>是<%else%>否<%end if%></td>
			<td class=td_rblue>腔数</td>
			<td class=td_lblue>
				<%if rs("qs")=1 then%>单腔<%end if%>
				<%if rs("qs")=2 then%>双腔<%end if%>
				<%if rs("qs")=3 then%>三腔<%end if%>
				<%if rs("qs")=4 then%>四腔<%end if%>
				<%if rs("qs")=5 then%>五腔<%end if%>
				<%if rs("qs")=6 then%>六腔<%end if%>
				<%if rs("qs")=7 then%>七腔<%end if%>
				<%if rs("qs")=8 then%>八腔<%end if%>
			</td>
		</tr>
		
		<tr>
			<td class=td_rblue>加热板信息</td>
			<td class=td_lblue>相数:<%=rs("jrbxs")%>	 材质:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
			<td class=td_rblue>牵引速度</td>
			<td class=td_lblue><%=rs("qysd")%> 米/分(m/min)</td>
		</tr>

		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>模具信息</b></td>
		</tr>

		<tr>
			<td class=td_rblue>模头结构</td>
			<td class=td_lblue><%=rs("mtjg")%>&nbsp;</td>
			<td class=td_rblue>任务内容</td>
			<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
		</tr>

		<tr>
			<td class=td_rblue>定型结构</td>
			<td class=td_lblue><%=rs("dxjg")%>&nbsp;</td>
			<td class=td_rblue>模头连接尺寸</td>
			<td class=td_lblue><%=rs("mtljcc")%></td>
		</tr>

		<tr>
			<td class=td_rblue>水箱结构</td>
			<td class=td_lblue><%=rs("sxjg")%>&nbsp;</td>
			<td class=td_rblue>热电偶规格</td>
			<td class=td_lblue><%=rs("rdogg")%></td>
		</tr>


		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>工艺信息</b></td>
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
			<td class=td_lblue height=25 colspan=4> <b>其他信息</b></td>
		</tr>

		<tr>
			<td class=td_rblue >评审记录</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("psjl"))%></td>
		</tr>

		<tr>
			<td class=td_rblue >备注</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("bz"))%></td>
		</tr>

		<tr>
			<td class=td_rblue>计划结束时间</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
			<td class=td_rblue>实际结束时间;</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
		</tr>

		<tr>
			<td class=td_rblue>组长</td>
			<td class=td_lblue><%=rs("zz")%></td>
			<td class=td_rblue>技术代表</td>
			<td class=td_lblue><%=rs("jsdb")%></td>
		</tr>
	</table>
<%
	response.write(hline(5, "100%", ""))
	call mtask_display_user(rs)
	response.write(hline(5, "100%", ""))
	call mtask_display_user2(rs)
end function

function mtask_display_user(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr>
		<%
			select case rs("mjxx") & rs("rwlr")
				case "全套设计"
				%>
					<%if (not isnull(rs("mtjgr"))) and (not isnull(rs("dxjgr"))) and (rs("mtjgr")=rs("dxjgr")) then%>
						<td class=td_blue width="10%" rowspan=2>模具结构</td>
					<%else%>
						<td class=td_blue width="10%">模头结构</td>
					<%end if%>
					<%if rs("mtjgr")=rs("dxjgr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtjgr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtjgr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="10%" rowspan=2>模具设计</td>
					<%else%>
						<td class=td_blue width="10%">模头设计</td>
					<%end if%>
					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtsjr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="10%" rowspan=2>模具审核</td>
					<%else%>
						<td class=td_blue width="10%">模头审核</td>
					<%end if%>
					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtshr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="10%" rowspan=2>模具BOM</td>
					<%else%>
						<td class=td_blue width="10%">模头BOM</td>
					<%end if%>
					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mtbomr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
						<td class=td_blue width="10%">定型结构</td>
					<%end if%>
					<%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
						<td class=td_blue width="15%"><%=rs("dxjgr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="10%">定型设计</td>
					<%end if%>
					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="10%">定型审核</td>
					<%end if%>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="10%">定型BOM</td>
					<%end if%>
					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
					<%end if%>
				<%
				case "全套复改"
				%>
					<td class=td_blue width="10%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>

					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="10%" rowspan=2>模具复改</td>
					<%else%>
						<td class=td_blue width="10%">模头复改</td>
					<%end if%>
					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtsjr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="10%" rowspan=2>模具审核</td>
					<%else%>
						<td class=td_blue width="10%">模头审核</td>
					<%end if%>
					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtshr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="10%" rowspan=2>模具BOM</td>
					<%else%>
						<td class=td_blue width="10%">模头BOM</td>
					<%end if%>
					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mtbomr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="10%">定型复改</td>
					<%end if%>
					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="10%">定型审核</td>
					<%end if%>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="10%">定型BOM</td>
					<%end if%>
					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
					<%end if%>
				<%
				case "全套复查"
				%>
					<td class=td_blue width="10%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="15%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="10%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="15%" rowspan=2>&nbsp;</td>

					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="10%" rowspan=2>模具复查</td>
					<%else%>
						<td class=td_blue width="10%">模头复查</td>
					<%end if%>
					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtshr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="10%" rowspan=2>模具BOM</td>
					<%else%>
						<td class=td_blue width="10%">模头BOM</td>
					<%end if%>
					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mtbomr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="10%">定型复查</td>
					<%end if%>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="10%">定型BOM</td>
					<%end if%>
					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
					<%end if%>
				<%
				case "模头设计"
				%>
					<td class=td_blue width="10%">模头结构</td>
					<td class=td_blue width="15%"><%=rs("mtjgr")%>&nbsp;</td>
					<td class=td_blue width="10%">模头设计</td>
					<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">模头审核</td>
					<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
				<%
				case "模头复改"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">模头复改</td>
					<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">模头审核</td>
					<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
				<%
				case "模头复查"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">模头复查</td>
					<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
				<%
				case "定型设计"
				%>
					<td class=td_blue width="10%">定型结构</td>
					<td class=td_blue width="15%"><%=rs("dxjgr")%>&nbsp;</td>
					<td class=td_blue width="10%">定型设计</td>
					<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">定型审核</td>
					<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
				<%
				case "定型复改"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">定型复改</td>
					<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">定型审核</td>
					<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
				<%
				case "定型复查"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">定型复查</td>
					<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
				<%
				case else
				response.write(rs("mjxx") & rs("rwlr"))
			end select
		%>
		</tr>
	</table>
<%
end function

function mtask_display_user2(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr>
		<%
			select case rs("mjxx")
				case "全套"
				%>
					<%if (not isnull(rs("mttsdr"))) and (not isnull(rs("dxtsdr"))) and (rs("mttsdr")=rs("dxtsdr")) then%>
						<td class=td_blue width="15%" rowspan=2>模具调试单</td>
					<%else%>
						<td class=td_blue width="15%">模头调试单</td>
					<%end if%>
					<%if rs("mttsdr")=rs("dxtsdr") then%>
						<td class=td_blue width="18%" rowspan=2><%=rs("mttsdr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="18%"><%=rs("mttsdr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mttsr")=rs("dxtsr") then%>
						<td class=td_blue width="15%" rowspan=2>模具调试</td>
					<%else%>
						<td class=td_blue width="15%">模头调试</td>
					<%end if%>
					<%if rs("mttsr")=rs("dxtsr") then%>
						<td class=td_blue width="18%" rowspan=2><%=rs("mttsr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="18%"><%=rs("mttsr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
						<td class=td_blue width="15%" rowspan=2>模具调试信息整理</td>
					<%else%>
						<td class=td_blue width="15%">模头调试信息整理</td>
					<%end if%>
					<%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mttsxxzlr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mttsxxzlr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
						<td class=td_blue width="15%">定型调试单</td>
					<%end if%>
					<%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
						<td class=td_blue width="18%"><%=rs("dxtsdr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
						<td class=td_blue width="15%">定型调试</td>
					<%end if%>
					<%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
						<td class=td_blue width="18%"><%=rs("dxtsr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
						<td class=td_blue width="15%">定型调试信息整理</td>
					<%end if%>
					<%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
						<td class=td_blue width="*"><%=rs("dxtsxxzlr")%>&nbsp;</td>
					<%end if%>
				<%
				case "模头"
				%>
					<td class=td_blue width="15%">模头调试单</td>
					<td class=td_blue width="18%"><%=rs("mttsdr")%>&nbsp;</td>
					<td class=td_blue width="15%">模头调试</td>
					<td class=td_blue width="18%"><%=rs("mttsr")%>&nbsp;</td>
					<td class=td_blue width="15%">模头调试信息整理</td>
					<td class=td_blue width="*"><%=rs("mttsxxzlr")%>&nbsp;</td>
				<%
				case "定型"
				%>
					<td class=td_blue width="15%">定型调试单</td>
					<td class=td_blue width="18%"><%=rs("dxtsdr")%>&nbsp;</td>
					<td class=td_blue width="15%">定型调试</td>
					<td class=td_blue width="18%"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=td_blue width="15%">定型调试信息整理</td>
					<td class=td_blue width="*"><%=rs("dxtsxxzlr")%>&nbsp;</td>
				<%
				case else
				response.write(rs("mjxx"))
			end select
		%>
		</tr>
	</table>
<%
end function


function mtask_display_user_all(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
	<%select case rs("mjxx") & rs("rwlr")%>
			<%case "全套设计"%>
				<tr>
					<td class=td_blue width="10%">模头结构</td>
					<td class=td_blue width="10%"><%=rs("mtjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头结构开始</td>
					<td class=td_blue width="20%"><%=rs("mtjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头结构结束</td>
					<td class=td_blue width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头设计</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头设计开始</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头设计结束</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头审核</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核开始</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核结束</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM开始</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM结束</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型结构</td>
					<td class=td_blue width="10%"><%=rs("dxjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型结构开始</td>
					<td class=td_blue width="20%"><%=rs("dxjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型结构结束</td>
					<td class=td_blue width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型设计</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型设计开始</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型设计结束</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型审核</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核开始</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核结束</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM开始</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM结束</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "模头设计"%>
				<tr>
					<td class=td_blue width="10%">模头结构</td>
					<td class=td_blue width="10%"><%=rs("mtjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头结构开始</td>
					<td class=td_blue width="20%"><%=rs("mtjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头结构结束</td>
					<td class=td_blue width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头设计</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头设计开始</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头设计结束</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头审核</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核开始</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核结束</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM开始</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM结束</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
			<%case "定型设计"%>
				<tr>
					<td class=td_blue width="10%">定型结构</td>
					<td class=td_blue width="10%"><%=rs("dxjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型结构开始</td>
					<td class=td_blue width="20%"><%=rs("dxjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型结构结束</td>
					<td class=td_blue width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型设计</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型设计开始</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型设计结束</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型审核</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核开始</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核结束</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM开始</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM结束</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "全套复改"%>
				<tr>
					<td class=td_blue width="10%">模头复改</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复改开始</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复改结束</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头审核</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核开始</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核结束</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM开始</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM结束</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型复改</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复改开始</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复改结束</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型审核</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核开始</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核结束</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM开始</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM结束</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "模头复改"%>
				<tr>
					<td class=td_blue width="10%">模头复改</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复改开始</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复改结束</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头审核</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核开始</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头审核结束</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM开始</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM结束</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
			<%case "定型复改"%>
				<tr>
					<td class=td_blue width="10%">定型复改</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复改开始</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复改结束</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型审核</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核开始</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型审核结束</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM开始</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM结束</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "全套复查"%>
				<tr>
					<td class=td_blue width="10%">模头复查</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复查开始</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复查结束</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM开始</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM结束</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型复查</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复查开始</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复查结束</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM开始</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM结束</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "模头复查"%>
				<tr>
					<td class=td_blue width="10%">模头复查</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复查开始</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头复查结束</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头BOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM开始</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头BOM结束</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
			<%case "定型复查"%>
				<tr>
					<td class=td_blue width="10%">定型复查</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复查开始</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型复查结束</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM开始</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型BOM结束</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
		<%end select%>
	</table>
<%
end function

function mtask_display_user2_all(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<%select case rs("mjxx")%>
			<%case "全套"%>
				<tr>
					<td class=td_blue width="10%">模头调试单</td>
					<td class=td_blue width="10%"><%=rs("mttsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试单开始</td>
					<td class=td_blue width="20%"><%=rs("mttsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试单结束</td>
					<td class=td_blue width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头调试</td>
					<td class=td_blue width="10%"><%=rs("mttsr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试开始</td>
					<td class=td_blue width="20%"><%=rs("mttsks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试结束</td>
					<td class=td_blue width="20%"><%=rs("mttsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头调试信息整理</td>
					<td class=td_blue width="10%"><%=rs("mttsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试信息整理开始</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试信息整理结束</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型调试单</td>
					<td class=td_blue width="10%"><%=rs("dxtsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试单开始</td>
					<td class=td_blue width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试单结束</td>
					<td class=td_blue width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型调试</td>
					<td class=td_blue width="10%"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试开始</td>
					<td class=td_blue width="20%"><%=rs("dxtsks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试结束</td>
					<td class=td_blue width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型调试信息整理</td>
					<td class=td_blue width="10%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试信息整理开始</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试信息整理结束</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
				</tr>
			<%case "模头"%>
				<tr>
					<td class=td_blue width="10%">模头调试单</td>
					<td class=td_blue width="10%"><%=rs("mttsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试单开始</td>
					<td class=td_blue width="20%"><%=rs("mttsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试单结束</td>
					<td class=td_blue width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头调试</td>
					<td class=td_blue width="10%"><%=rs("mttsr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试开始</td>
					<td class=td_blue width="20%"><%=rs("mttsks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试结束</td>
					<td class=td_blue width="20%"><%=rs("mttsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">模头调试信息整理</td>
					<td class=td_blue width="10%"><%=rs("mttsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试信息整理开始</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">模头调试信息整理结束</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
				</tr>
			<%case "定型"%>
				<tr>
					<td class=td_blue width="10%">定型调试单</td>
					<td class=td_blue width="10%"><%=rs("dxtsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试单开始</td>
					<td class=td_blue width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试单结束</td>
					<td class=td_blue width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型调试</td>
					<td class=td_blue width="10%"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试开始</td>
					<td class=td_blue width="20%"><%=rs("dxtsks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试结束</td>
					<td class=td_blue width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">定型调试信息整理</td>
					<td class=td_blue width="10%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试信息整理开始</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">定型调试信息整理结束</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
				</tr>

			<%
				case else
				response.write(rs("mjxx"))
			end select
			%>
		</tr>
	</table>
<%
end function
%>