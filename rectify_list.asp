<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="问题分析 → 纠正/预防措施表"
strPage="tech"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Dim strFeedBack,strbh,strzrbm,strxxbm,strjssj,strbhgnr,strwczk
strbh=Request("bh")
strzrbm=Request("zrbm")
strxxbm=Request("xxbm")
strjssj=Request("jssj")
strbhgnr=Request("bhgnr")
strwczk=Request("wczk")
strFeedBack="&zrbm="&strzrbm&"&xxbm="&strxxbm&"&bhgnr="&strbhgnr

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
			<%Call RectifyList()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function SearchInfo()
%>
	<Table border=0 cellpadding=2 cellspacing=0 width="100%">
		<Form action="<%=Request.Servervariables("script_name")%>" method=get>
		<Tr><Td>
		&nbsp;&nbsp;选择筛选条件:
			编号:<input type="text" name="bh" size=8 value="<%=strbh%>">
			责任部门:<input type="text" name="zrbm" size=8 value="<%=strzrbm%>">
			不合格/潜在不合格内容:<input type="text" name="bhgnr" size=30 value="<%=strbhgnr%>">
		<input type=submit value=" 选 择 ">
		</Td></Tr>
		</Form>
	</Table>
<%
End Function

Function RectifyList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount
	absPageNum = 0
	RecordPerPage = 20
	strSql=""
	If strbh<>"" Then strSql=" bh='" & strbh & "'" End If
	If strzrbm<>"" Then
		strSql=strSql & " zrbm='"&strzrbm&"' "
	End If
	If strbhgnr<>"" Then
		If strSql<>"" Then
			strSql=strSql & " and bhgnr like '%" & strbhgnr &"%'"
		Else
			strSql=" bhgnr like '%" & strbhgnr &"%'"
		End IF
	End If

	If strSql<>"" Then
		strSql="select * from [Rectify]  where" & strSql & " order by jssj desc"
	Else
		strSql="select * from [Rectify] order by jssj desc"
	End If

	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,3,3
	If Rs.eof Or Rs.bof Then
		Call TbTopic("没有任何内容符合条件！") : Exit Function
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
	icounter=totalcount-(abspagenum-1)*recordperpage
	Call TbTopic("纠正/预防措施表")

%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable  align="center">
		<tr>
			<th class=th width=40>ID</th>
			<th class=th width=40>编号</th>
			<th class=th width=60>责任部门</th>
			<th class=th width=80>发出信息部门</th>
			<th class=th width=*>不合格/潜在不合格内容</th>
			<th class=th width=100>信息发出日期</th>
			<th class=th width=60>当前状态</th>
		</tr>
<%
	for absrecordnum = 1 to recordperpage
%>

		<tr>
			<td class=ctd><%=icounter%></td>
			<td class=ctd alt="点击查看具体信息"><a href="Rectify_dis.asp?s_lsh=<%=rs("bh")%>&id=<%=rs("id")%>"><%=rs("bh")%></a></td>
			<td class=ctd><%=rs("zrbm")%></td>
			<td class=ctd><%=rs("xxbm")%></td>
			<td class=ltd><%= xjweb.StringCut(rs("bhgnr"),26)%></td>
			<td class=ctd><%=rs("jssj")%></td>
			<td class=ctd><%=rs("wczk")%></td>
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