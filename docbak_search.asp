<!--#include file="include/conn.asp"-->
<%
'10:06 2007-1-30-星期二
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="图档备份 → 查询存档信息"
strPage="docbak"
'Call FileInc(0, "js/docbak.js")
xjweb.header()
Call TopTable()

Dim strddh, strlsh, strdwmc, strdiskid, strmh, strFeedBack
strddh = trim(request("ddh"))
strlsh = trim(request("lsh"))
strdwmc = trim(request("dwmc"))
strdiskid = trim(request("diskid"))
strmh = trim(request("mh"))
strFeedBack="&lsh="&strLsh&"&ddh="&strDdh&"&dwmc="&strDwmc&"&diskid="&strdiskid&"&mh="&strmh

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
			<%Call DocBakSearch()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function SearchInfo()
%>
	<table border=0 width="100%" height="100%" cellpadding=2 cellspacing=0>
		<form action="<%=request.servervariables("script_name")%>" method="get">
		<tr>
			<td height=25>
				&nbsp;输入筛选条件:
				订单号:<input type="text" name="ddh" value="<%=strddh%>" size="8">&nbsp;
				流水号:<input type="text" name="lsh" value="<%=strlsh%>" size="6">&nbsp;
				单位名称:<input type="text" name="dwmc" value="<%=strdwmc%>" size="8">&nbsp;
				所存盘号:<input type="text" name="diskid" value="<%=strdiskid%>" size="6">&nbsp;
				模号:<input type="text" name="mh" value="<%=strmh%>" size="10">&nbsp;
				<input type="submit" value=" 提取 ">
			</td>
		</tr>
		</form>
	</table>
<%
End Function

Function DocBakSearch()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount, strTerm
	absPageNum = 0
	RecordPerPage = 20
	strSql=""
	if strddh <> "" then strSql = " ddh like '%"&strddh&"%'"
	if strlsh <> "" then
		if strSql <> "" then
			strSql = strSql & " and lsh like '%"&strlsh&"%'"
		else
			strSql = " lsh like '%"&strlsh&"%'"
		end if
	end if

	if strdwmc <> "" then
		if strSql <> "" then
			strSql = strSql & " and dwmc like '%"&strdwmc&"%'"
		else
			strSql = " dwmc like '%"&strdwmc&"%'"
		end if
	end if

	if strdiskid <> "" then
		if strSql <> "" then
			strSql = strSql & " and diskid like '%"&strdiskid&"%'"
		else
			strSql = " diskid like '%"&strdiskid&"%'"
		end if
	end if

	if strmh <> "" then
		if strSql <> "" then
			strSql = strSql & " and mh like '%"&strmh&"%'"
		else
			strSql = " mh like '%"&strmh&"%'"
		end if
	end if

	If strSql <> "" Then
		strSql="select * from [doc_bak] where " & strSql & " order by cpsj desc"
	Else
		strSql="select * from [doc_bak] order by cpsj desc"
	End If

	'Response.write strSql
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,1,3
	If Rs.eof Or Rs.bof Then
		Call JsAlert("没有任何存档信息，请放宽筛选条件！","docbak_search.asp") : Exit Function
	End If
	Rs.PageSize = RecordPerPage
	TotalCount=Rs.RecordCount

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

	If absPageNum > Rs.PageCount then absPageNum = Rs.PageCount
	rs.absolutepage = absPageNum
	iCounter=totalcount-(abspagenum-1)*recordperpage
	Call TbTopic("挤出模具厂存档信息列表")

	'icounter = (absPageNum - 1) * recordperpage + 1
%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable  align="center">
		<tr>
			<th class=th>id</th>
			<th class=th>订单号</th>
			<th class=th>流水号</th>
			<th class=th>单位名称</th>
			<th class=th>模号</th>
			<th class=th>所存盘号</th>
			<th class=th>备注</th>
			<%if chkable(7) then response.write("<th class=th>操作</th>")%>
		</tr>
<%
	for absrecordnum = 1 to recordperpage
%>

		<tr>
			<td class=ctd><%=iCounter%></td>
			<td class=ctd><%=rs("ddh")%>&nbsp;</td>
			<td class=ctd alt="流 水 号: <%=rs("lsh")%><br>模具设计: <%=rs("mjsj")%><br>模具审核: <%=rs("mjsh")%><br>工艺设计: <%=rs("gysj")%><br>工艺审核: <%=rs("gysh")%>"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a>&nbsp;</td>
			<td class=ctd><%=rs("dwmc")%></td>
			<td class=ctd><%=ucase(rs("mh"))%>&nbsp;</td>
			<td class=ctd alt="存盘时间:<%=rs("cpsj")%>"><%=rs("diskid")%></td>
			<td class=ctd>
			<%if trim(rs("bz")) <> "" then%>
				<div alt="备注:<br><%=rs("bz")%>">备注</div>
			<%else%>
				&nbsp;
			<%end if%>
			</td>

			<%if chkable(7) and not isnull(rs("lsh")) then%>
				<%If not isnull(rs("lsh")) Then%>
					<td class=ctd><a href="docbak_change.asp?s_lsh=<%=rs("lsh")%>">更改</a></td>
				<%else%>
					<td class=ctd>&nbsp;</td>
				<%End If%>
			<%end if%>
		</tr>
		<%rs.movenext%>
		<%if rs.eof then%>
			<%exit for%>
		<%end if%>
		<%icounter = icounter - 1%>
	<%next%>
	</table>
	<table width="95%" cellpadding=2 cellspacing=0 border=0 align="center">
		<tr>
			<td align=left>
				符合条件共 <%=rs.recordcount%> 个&nbsp;&nbsp;
				每页 <%=rs.pagesize%> 个&nbsp;&nbsp;
				共 <%=Rs.PageCount%> 页&nbsp;&nbsp;
				当前为第 <%=absPageNum%> 页
			</td>
			<td align=right>
				【
			<%
				if absPageNum > 1 then
					response.write("<a href="&Request.ServerVariables("script_name")&"?ipage="&(abspagenum-1)&strFeedBack&" alt='上一页'> ←</a>&nbsp;&nbsp;")
				end if
				Dim iStart,iEnd
				if absPageNum < 4 then
					iStart = 1
				else
					iStart = absPageNum - 3
				end if
				if absPageNum < Rs.PageCount - 3 then
					iEnd = absPageNum + 3
				else
					iEnd = Rs.PageCount
				end if
				for i = iStart to iEnd
					if i = absPageNum then
						response.write("&nbsp;<font style=font-size:11pt;><b>"&  i & "</b></font>&nbsp;")
					else
						response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&i&strFeedBack&">" & i & "</a>&nbsp;")
					end if
				next
				if absPageNum < Rs.PageCount then
					response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&(absPageNum+1)&strFeedBack&" alt='下一页'> → </a>&nbsp;")
				end if
			%>
				】
				跳转到:
				<select name="ipage" onchange='location.href("<%=Request.ServerVariables("script_name")%>?ipage=" + this.value+"<%=strFeedBack%>");'>
					<%for i=1 to Rs.PageCount%>
						<%if i = absPageNum then%>
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