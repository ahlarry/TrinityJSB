<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="调试任务 → 调试任务列表"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="atask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

Dim strFeedBack, strlsh, strddh, strdwmc, strDmmc, strZuz, strorder, strterm
strlsh = trim(request("lsh"))
strddh = trim(request("ddh"))
strdwmc = trim(request("dwmc"))
strDmmc=Trim(Request("dmmc"))
strZuz=Trim(Request("zuz"))
strorder = trim(request("order"))
strterm = trim(request("term"))

strFeedBack = "&lsh="&strlsh&"&ddh="&strddh&"&dwmc="&strdwmc&"&dmmc="&strDmmc&"&zuz="&strZuz&"&order="&strorder&"&term="&strterm

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchInfo()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call ataskList()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function SearchInfo()
%>
	<table width="100%" cellpadding=2 cellspacing="0" border="0" height="100%">
		<form action="<%=request.servervariables("script_name")%>" method="get">
		<tr>
			<td height="25">
				&nbsp;&nbsp;输入筛选条件:
				订单号:<input type="text" name="ddh" value="<%=strddh%>" size="12">
				流水号:<input type="text" name="lsh" value="<%=strlsh%>" size="6">
				单位名称:<input type="text" name="dwmc" value="<%=strdwmc%>" size="8">
				断面名称:<input type="text" name="dmmc" value="<%=strDmmc%>" size="8">
				组长:<input type="text" name="zuz" value="<%=strZuz%>" size="8">
				<input type="submit" value=" 查找 ">
				<p>&nbsp;&nbsp;排序:
				<select name="order" onchange='location.href("<%=request.servervariables("script_name")%>?ipage=1&lsh=<%= strlsh%>&ddh=<%=strddh%>&dwmc=<%=strdwmc%>&term=<%=strterm%>&order=" + this.value);'>
					<%If strorder = "ddh" Then%>
						<option value="ddh" selected>订单号</option>
						<option value="lsh">流水号</option>
					<%Else%>
						<option value="ddh">订单号</option>
						<option value="lsh" selected>流水号</option>
					<%End If%>
				</select>
				&nbsp;&nbsp;
				条件:
				<select name="term" onchange='location.href("<%=request.servervariables("script_name")%>?ipage=1&lsh=<%= strlsh%>&ddh=<%=strddh%>&dwmc=<%=strdwmc%>&order=<%=strorder%>&term=" + this.value);'>
					<%select case strterm%>
						<%case "no"%>
							<option value="no" selected>未完成</option>
							<option value="ok">已完成</option>
							<option value="all">全部</option>
						<%case "ok"%>
							<option value="no">未完成</option>
							<option value="ok" selected>已完成</option>
							<option value="all">全部</option>
						<%case "all"%>
							<option value="no">未完成</option>
							<option value="ok">已完成</option>
							<option value="all" selected>全部</option>
						<%case else%>
							<option value="no">未完成</option>
							<option value="ok">已完成</option>
							<option value="all" selected>全部</option>
					<%end select%>
				</select>
			</p></td>
		</tr>
		</form>
	</table>
<%
End Function

Function ataskList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter
	absPageNum = 0
	strSql = ""
	RecordPerPage = 20
	if strlsh <> "" then
		strsql = " lsh like '%"&strlsh&"%' "
	end if
	if strdwmc <> "" then
		if strsql <> "" then
			strsql = " dwmc like '%"&strdwmc&"%' and " & strsql
		else
			strsql  = " dwmc like '%"&strdwmc&"%' "
		end if
	end if
	If strDmmc <> "" Then
		If strSql <> "" Then
			strSql = " dmmc like '%"&strDmmc&"%' and " & strSql
		Else
			strSql  = " dmmc like '%"&strDmmc&"%' "
		End If
	End If
	If strZuz <> "" Then
			If strSql <> "" Then
			strSql = " zz like '%"&strZuz&"%' or jgzz like '%"&strZuz&"%' or sjzz like '%"&strZuz&"%' and " & strSql
		Else
			strSql  = " zz like '%"&strZuz&"%' or jgzz like '%"&strZuz&"%' or sjzz like '%"&strZuz&"%' "
		End If
	End If
	if strddh <> "" then
		if strsql <> "" then
			strsql = " ddh like '%"&strddh&"%' and " & strsql
		else
			strsql  = " ddh like '%"&strddh&"%' "
		end if
	end if

	select case strterm
			case "no"
				if strsql <> "" then
					strsql = " not(mjjs) and " & strsql
				else
					strsql  = " not(mjjs) "
				end if
			case "ok"
				if strsql <> "" then
					strsql = " mjjs and " & strsql
				else
					strsql  = " mjjs "
				end if
			case "all"

			case else
	end select

	if request("strsql") <> "" then
		strsql = request("strsql")
	end if

	Dim sqlorder, Tmplsh
	sqlorder = " order by lsh desc"
	If LCase(strorder) = "ddh" Then sqlorder = " order by ddh desc, lsh desc"
	If LCase(strorder) = "lsh" Then sqlorder = " order by lsh desc"
	If strsql <> "" Then
		strsql = "select * from [mtask] where (not(isnull(sjjssj)) and datediff('m',sjjssj,'"&now()&"')<15 and " & strsql & ") or (rwlr='修理' and datediff('m',jhjssj,'"&now()&"')<12)" & sqlorder
	Else
		strsql = "select * from [mtask] where (not(isnull(sjjssj)) and datediff('m',sjjssj,'"&now()&"')<15) or (rwlr='修理' and datediff('m',jhjssj,'"&now()&"')<12)" & sqlorder
	End If
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize = RecordPerPage
	Rs.Open strSql,Conn,1,3
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("指定的条件没有任何任务书，请重新选择条件！","atask.asp")
	End If

	Rs.PageSize = RecordPerPage

	If Trim(Request("iPage")) <> ""  Then
		If IsNumeric(Trim(Request("iPage"))) Then
			If Trim(Request("iPage")) <= 0 Then
				absPageNum = 1
			ElseIf CLng(Trim(Request("iPage"))) > Rs.PageCount Then
				absPageNum = Rs.PageCount
			Else
				absPageNum = CLng(Trim(Request("iPage")))
			End If
		Else
			If Request("iCurPage") <> "" Then
				absPageNum = CLng(Request("iCurPage"))
			Else
				absPageNum = 1
			End If
		End If
	Else
		If Request("iCurPage") <> "" Then
			absPageNum = CLng(Request("iCurPage"))
		Else
			absPageNum = 1
		End If
	End If

	If Trim(Request("iPage")) <> "" And IsNumeric(Request("iPage")) Then
		absPageNum=CLng(Request("iPage"))
	End If

	If absPageNum>Rs.PageCount Then absPageNum=Rs.PageCount
	Rs.absolutePage = absPageNum
	iCounter=(absPageNum-1)*RecordPerPage+1
	Call CutLine()		'显示图例
	Call TbTopic("挤出模具厂辅助任务列表")
%>
	<Table width="98%" cellpadding=2 cellspacing=0 border=0  class=xtable align="center">
		<tr>
			<th class=th>id</th>
			<th class=th>订单号</th>
			<th class=th>流水号</th>
			<th class=th>单位名称</th>
			<th class=th>断面名称</th>
			<th class=th>组长</th>
			<th class=th>任务内容</th>
			<th class=th width=*>调试单</th>
			<th class=th width=*>调试</th>
			<th class=th width=*>调试整理</th>
			<th class=th width=*>齐套整理</th>
		</tr>
<%
	For absRecordNum=1 To RecordPerPage
	If rs("rwlr")="修理" Then
		Tmplsh=rs("lsh")&"["&rs("mh")&"]"
	Else
		Tmplsh=rs("lsh")
	End If
%>
		<tr>
			<td class=ctd><%=iCounter%></td>
			<td class=ctd><a href="?ddh=<%=rs("ddh")%>"><%=rs("ddh")%></a></td>
			<td class=ctd alt="<%=rs("bm")%>"><a href=atask_display.asp?s_lsh=<%=rs("lsh")%>>
				<%If InStr(rs("bz"),"调试关注")>0 Then
      		 		Response.Write("<font color=red><b>"&Tmplsh&"</b></font>")
      			Else
      				Response.Write(Tmplsh)
      			End If%>
      		</a></td>
			<td class=ctd><%=rs("dwmc")%></td>
			<td class=ctd alt="断面名称: <%=rs("dmmc")%>"><%=xjweb.StringCut(rs("dmmc"),12)%></td>
			<td class=ctd><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(j)、"&rs("sjzz")&"(s)")%></td>
			<td class=ctd><%=rs("mjxx") &  rs("rwlr")%></td>
		<%select case rs("mjxx")%>
			<%case "全套"%>
					<td class=ctd>
						<%call distd(rs("mttsdks"),rs("mttsdjs"),-20,rs)%>
						<%call distd(rs("dxtsdks"),rs("dxtsdjs"),-20,rs)%>
					</td>

					<td class=ctd>
						<%call distd2(rs("mttsks"),rs("mttsjs"),rs)%>
						<%call distd2(rs("dxtsks"),rs("dxtsjs"),rs)%>
					</td>
					<td class=ctd>
						<%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
						<%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
					</td>
					<td class=ctd>
						<%=Rs("xtxxjhjs")%>
						<%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%>
					</td>
			<%case "模头"%>
					<td class=ctd>
						<%call distd(rs("mttsdks"),rs("mttsdjs"),-20,rs)%>
					</td>

					<td class=ctd>
						<%call distd2(rs("mttsks"),rs("mttsjs"),rs)%>
					</td>

					<td class=ctd>
						<%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
					</td>
					<td class=ctd>
						<%=Rs("xtxxjhjs")%>
						<%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%>
					</td>
			<%case "定型"%>
					<td class=ctd>
						<%call distd(rs("dxtsdks"),rs("dxtsdjs"),-20,rs)%>
					</td>

					<td class=ctd>
						<%call distd2(rs("dxtsks"),rs("dxtsjs"),rs)%>
					</td>

					<td class=ctd>
						<%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
					</td>
					<td class=ctd>
						<%=Rs("xtxxjhjs")%>
						<%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%>
					</td>
		<%end select%>
		</tr>
		<%rs.movenext%>
		<%if rs.eof then%>
			<%exit for%>
		<%end if%>
		<%icounter = icounter + 1%>
	<%next%>
	</table>
	<table width="95%" cellpadding=2 cellspacing=0 border=0 align="center">
		<tr>
			<td align=left>
				符合条件共 <%=rs.recordcount%> 个&nbsp;&nbsp;
				每页 <%=rs.pagesize%> 个&nbsp;&nbsp;
				共 <%=rs.pagecount%> 页&nbsp;&nbsp;
				当前为第 <%=abspagenum%> 页
			</td>
			<td align=right>
				【
			<%
				If absPageNum > 1 Then
					Response.Write("<a href="&request.servervariables("script_name")&"?ipage="&(abspagenum-1)&strFeedBack&" title='上一页'> ←</a>&nbsp;&nbsp;")
				end if
				Dim iStart,iEnd
				If absPageNum<4 Then
					iStart=1
				Else
					iStart = absPageNum-3
				End If
				If absPageNum < rs.PageCount - 3 Then
					iEnd = absPageNum + 3
				Else
					iEnd = rs.PageCount
				End If
				For i = iStart To iEnd
					if i = abspagenum then
						response.write("&nbsp;<font style=font-size:11pt;><b>"&  i & "</b></font>&nbsp;")
					else
						response.write("&nbsp;<a href="&request.servervariables("script_name")&"?ipage="&i&strFeedBack&">" & i & "</a>&nbsp;")
					end if
				next
				if abspagenum < rs.pagecount then
					response.write("&nbsp;<a href="&request.servervariables("script_name")&"?ipage="&(abspagenum+1)&strFeedBack&" title='下一页'> → </a>&nbsp;")
				end if
			%>
				】
				跳转到:
				<select name="ipage" onchange='location.href("<%=request.servervariables("script_name")%>?ipage=" + this.value +"<%=strFeedback%>");'>
					<%for i=1 to rs.pagecount%>
						<%if i = abspagenum then%>
							<option value=<%=i%> selected>第 <%=i%> 页</option>
						<%else%>
							<option value=<%=i%>>第 <%=i%> 页</option>
						<%end if%>
					<%next%>
				</select>
				</td>
			</tr>
		</table>
<%
end function
%>