<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="员工考评 → 员工考评列表"
strPage="ygkp"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

Dim strFeedBack,strZrr, strkpitem, strgzz, strkpjs, strclsh, striPage
striPage=NullToNum(Request("ipage"))
strZrr=Trim(Request("zrr"))
strkpitem = trim(request("kpitem"))
strkpjs = trim(request("kpjs"))
strclsh = trim(request("clsh"))
strgzz =request("gzz")
If strgzz="" Then strgzz=0
strFeedBack=""
If strZrr<>"" Then strFeedBack="&zrr="&strZrr
If strkpitem<>"" Then strFeedBack="&kpitem="&strkpitem&strFeedBack
If strgzz<>"0" Then strFeedBack="&gzz="&strgzz&strFeedBack
If strkpjs<>"" Then strFeedBack="&kpjs="&strkpjs&strFeedBack
If strclsh<>"" Then strFeedBack="clsh="&strclsh&"&"&strFeedBack

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
			<%Call ygkpList()%>
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
		&nbsp;&nbsp;
			责任人:
			<Select name="zrr" onchange="javascript:window.location.href('<%=request.servervariables("script_name")%>?zrr=' + this.value+'&kpitem='+this.form.kpitem.value+'&kpjs='+this.form.kpjs.value);">
				<option value="">全部</option>
				<%If Session("userName")<>"" Then%>
					<option value="<%=Session("userName")%>"><%=Session("userName")%></option>
				<%End If%>
				<%If ChkAble("1,2,3,4,10,11") Then%>
					<%for i = 0 to ubound(c_allzy)%>
						<option value="<%=c_allzy(i)%>"<%If strZrr=c_allzy(i) Then%> Selected<%End If%>><%=c_allzy(i)%></option>
					<%next%>
				<%End If%>
			</Select>&nbsp;&nbsp;&nbsp;
			工作组：
				<Select name="gzz" onchange="javascript:window.location.href('<%=request.servervariables("script_name")%>?gzz=' + this.value+'&kpitem='+this.form.kpitem.value+'&kpjs='+this.form.kpjs.value);">
				<option value=0>全部</option>
				<%If ChkAble("1,2,3,4") Then%>
					<%for i = 1 to 5%>
						<option value=<%=i%> <%If Cint(strgzz)=i Then%> Selected<%End If%>><%=i%></option>
					<%next%>
				<%End If%>
			</Select>&nbsp;&nbsp;&nbsp;
			考评项目:&nbsp;<input type="text" name="kpitem" value="<%=strkpitem%>" size="10">&nbsp;&nbsp;
			角色:&nbsp;<input type="text" name="kpjs" value="<%=strkpjs%>" size="8">&nbsp;&nbsp;
			流水号:&nbsp;<input type="text" name="clsh" value="<%=strclsh%>" size="8">&nbsp;&nbsp;
		<input type=submit value=" 选 择 ">
		</Td></Tr>
		</Form>
	</Table>
<%
End Function

Function ygkpList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount
	absPageNum = 0
	RecordPerPage = 20
	strSql=""
	if strkpitem <> "" then
			strSql = " kp_item like '%"&strkpitem&"%'"
	end if
	If strkpjs<>"" Then
		If strSql<>"" Then
			strSql=" kp_zrrjs='"&strkpjs&"' and " & strSql
		Else
			strSql=" kp_zrrjs='"&strkpjs&"' "
		End If
	End If
	If strgzz<>0 Then
		If strSql<>"" Then
			strSql=" kp_group="&cstr(strgzz)&" and " & strSql
		Else
			strSql=" kp_group="&cstr(strgzz)
		End If
	End If
	If strZrr<>"" Then
		If strSql<>"" Then
			strSql=" kp_zrr='"&strZrr&"' and " & strSql
		Else
			strSql=" kp_zrr='"&strZrr&"' "
		End If
	End If
	If strclsh<>"" Then
		If strSql<>"" Then
			strSql=" kp_lsh='"&strclsh&" or kp_bz like '%"&strclsh&"%' and " & strSql
		Else
			strSql=" kp_lsh='"&strclsh&"' or kp_bz like '%"&strclsh&"%'"
		End If
	End If
	If strSql<>"" Then
		strSql="select * from [kp_jsb]  where" & strSql & " order by kp_time desc"
	Else
		strSql="select * from [kp_jsb] order by kp_time desc"
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
	Call TbTopic("挤出模具厂员工考评列表")
	'icounter = (absPageNum - 1) * recordperpage + 1
%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable  align="center">
		<tr>
			<th class=th width=40>ID</th>
			<th class=th width=*>考评项目</th>
			<th class=th width=60>考评分</th>
			<th class=th width=60>因子</th>
			<th class=th width=60>责任人</th>
			<th class=th width=80>角色</th>
			<th class=th width=100>考评时间</th>
			<%If ChkAble("1,2,3,11,12") Then%>
				<th class=th width=100>操作</th>
			<%End If%>
		</tr>
<%
	for absrecordnum = 1 to recordperpage
%>

		<tr onmouseover="bgColor='#EEEEEE';" onmouseout="bgColor='';">
			<td class=ctd><%=icounter%></td>
			<td class=ltd><span alt="<%=replace(replace(replace(Rs("kp_bz"),vbcrlf,"<br>"),"'","&#39;"),Chr(34),"&#34;") & "<br><br><b>考评人</b>:" & Rs("kp_kpr")%>"><%=Rs("kp_item")%>&nbsp;</span></td>
			<td class=ctd><%=Formatnumber(Rs("kp_uprice"),1,-1)%></td>
			<td class=ctd><%=Rs("kp_mul")%>&nbsp;</td>
			<td class=ctd><%=Rs("kp_zrr")%></td>
			<td class=ctd><%=Rs("kp_zrrjs")%>&nbsp;</td>
			<td class=ctd><%=xjDate(rs("kp_time"),1)%></td>
			<%If ChkAble("1,2,3,11,12") Then%>
			<td class=ctd>
				&nbsp;
				<%If Rs("kp_kpr")=Session("userName") or ChkAble("3") Then%><a href="ygkp_change.asp?id=<%=Rs("ID")%>&kind=<%=Rs("kp_kind")%>&group=<%=Rs("kp_group")%>&ipage=<%=striPage&strFeedBack%>">更改</a><%End If%>
				<%If ChkAble(3) or (Rs("kp_kpr")=Session("userName") and Session("userdepart")="品管部") Then%>
				 | <a href="ygkp_delete.asp?id=<%=Rs("ID")%>&ipage=<%=striPage&strFeedBack%>" onclick="return confirm('删除后将不能恢复!\n\n确认删除 第 <%=iCounter%> 条 考评信息吗?');">删除</a><%End If%>
				&nbsp;
			</td>
			<%End If%>
		</tr>
		<%rs.movenext%>
		<%if rs.eof then%>
			<%exit for%>
		<%end if%>
		<%icounter = icounter - 1%>
	<%next%>
	</table>
	<table width="95%" cellpadding=2 cellspacing=0 border=0  align="center">
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