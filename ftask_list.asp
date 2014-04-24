<!--#include file="include/conn.asp"-->
<!--#include file="include/page/ftask_dbinf.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="零星任务 → 零星任务列表"
strPage="ftask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

Dim strFeedBack,strTerm,strType,strxldh,strxlxh,strxlcj,ifz
strTerm=Trim(Request("term"))
strType=Trim(Request("type"))
strxldh=Trim(Request("xldh"))
strxlxh=Trim(Request("xlxh"))
strxlcj=Trim(Request("xlcj"))
ifz=(Request("fz"))
strFeedBack="&term="&strTerm&"&type="&strType&"&xldh="&strxldh&"&xlxh="&strXlxh&"&xlcj="&strXlcj&"&fz="&ifz

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchInfo()%>
    </td>
  </tr>
  <Tr>
    <Td class=ctd height=300><%Call ftaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</Table>
<%
End Sub
Function Task_Ts(mystr)
       dim strnew
		if instr(mystr,"||")>0then
		strnew=split(mystr,"||")
	    for i=0 to ubound(strnew)
		Task_Ts=Task_Ts&strnew(i)&"<br>"
		next
		else
		Task_Ts=mystr
		end if

End Function


Function SearchInfo()
%>
<table>
  <tr>
    <td></td>
    <td></td>
  </tr>
</table>
<Table border=0 cellpadding=2 cellspacing=0 width="100%">
  <Form action="<%=Request.Servervariables("script_name")%>" method=get>
    <Tr>
      <Td>&nbsp;&nbsp;<strong>查找&nbsp;》》</strong>&nbsp;修理单号:</td>
      <Td><input type="text" name="xldh" size="12" value=<%=strxldh%>></td>
      <Td>修理小号:
        <input type="text" name="xlxh" size="10" value=<%=strxlxh%>>
      </td>
      <Td> 修理厂家:
        <input type="text" name="xlcj" size="10" value=<%=strxlcj%>></td>
      <Td> 分值:
        <input type="text" name="fz" size="5" value=<%=ifz%>></td>
      <td rowspan="2"><input type=submit value=" 选择 ">
      </Td>
    </Tr>
    <tr>
      <td>&nbsp;&nbsp;<strong>筛选&nbsp;》》</strong>&nbsp;&nbsp;责任人:</td>
      <Td><Select name="term" onchange='window.location.href("<%=request.servervariables("script_name")%>?type=<%=strType%>&xldh=<%=strxldh%>&xlxh=<%=strxlxh%>&xlcj=<%=strxlcj%>&fz=<%=ifz%>&term=" + this.value);'>
          <option value="">全部</option>
          <%If Session("userName")<>"" Then %>
          <option value="<%=Session("userName")%>"><%=Session("userName")%></option>
          <%End If%>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>' <%If strTerm=c_jsb(i) Then%> Selected<%End If%>><%=c_jsb(i)%></option>
          <%next%>
        </Select>
      </td>
      <Td> 任务类型:
        <select name="type" onchange='location.href("<%=request.servervariables("script_name")%>?term=<%=strterm%>&xldh=<%=strxldh%>&xlxh=<%=strxlxh%>&xlcj=<%=strxlcj%>&fz=<%=ifz%>&type=" + this.value);'>
          <option value="">全部</option>
          <%for i = 0 to ubound(c_lxrwlx)%>
          <option value='<%=c_lxrwlx(i)%>'<%If strType=c_lxrwlx(i) Then%> Selected<%End If%>><%=c_lxrwlx(i)%></option>
          <%next%>
        </select>
      </td>
      <td colspan="2">&nbsp;&nbsp;</td>
    </tr>
  </Form>
</Table>
<%
End Function

Function ftaskList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount, strTerm, strType, strSql,strxldh,strXlxh,strXlcj,ifz
	absPageNum = 0
	RecordPerPage = 20
	strSql = "" : strxldh="":ifz=""
	strTerm=Trim(Request("term"))
	strType=Trim(Request("type"))
	strxldh=Trim(Request("xldh"))
	strXlxh=Trim(Request("xlxh"))
	strXlcj=Trim(Request("Xlcj"))
	ifz=(Request("fz"))

	If strTerm <> "" Then
		strSql = " zrr='"&strTerm&"' "
	End If

	If strType <> "" Then
		If strSql <> "" Then
			strSql = "rwlx like '%"&strType&"%' and " & strSql
		Else
			strSql  = "rwlx like '%"&strType&"%' "
		End If
	End If

	If strxldh <> "" Then
		If strSql <> "" Then
			strSql = "xldh='"&strxldh&"' and " & strSql
		Else
			strSql  = "xldh='"&strxldh&"' "
		End If
	End If

	If strXlxh <> "" Then
		If strSql <> "" Then
			strSql = "rwlr like '%"&strxlxh&"%' and " & strSql
		Else
			strSql  = "rwlr like '%"&strxlxh&"%' "
		End If
	End If

	If strXlcj <> "" Then
		If strSql <> "" Then
			strSql = "rwlr like '%"&strXlcj&"%' and " & strSql
		Else
			strSql  = "rwlr like '%"&strXlcj&"%' "
		End If
	End If

	If ifz<>"" Then
		If strSql <> "" Then
			strSql = "zf="&ifz&" and " & strSql
		Else
			strSql  = "zf="&ifz&""
	    End If
	End If

	If strSql <> "" Then
		strSql = "select * from [ftask] where "&strSql & " order by jssj desc"
	Else
		strSql = "select * from [ftask] order by jssj desc"
	End If
	Set Rs = Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Call xjweb.Exec("",-1)
	Rs.open strSql,Conn,3,3
	If Rs.Eof Or Rs.Bof Then
		Call TbTopic(strTerm & " 没有任何"&strType&"任务！") : Exit Function
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
	Call TbTopic("挤出模具厂零星任务列表")

	'icounter = (absPageNum - 1) * recordperpage + 1
%>
<table width="95%" cellpadding=2 cellspacing=0 class=xtable align="center">
  <tr>
    <th class=th width=40>ID</th>
    <th class=th width=*>任务类型</th>
    <th class=th width=*>完成日期</th>
    <th class=th width=100>分值</th>
    <th class=th width=100>责任人</th>
    <%if chkable(3) then%>
    <th class=th width=120 colspan=2>操  作</th>
    <%end if%>
  </tr>
  <%
	for absrecordnum = 1 to recordperpage
	call Task_Ts(rs("rwlr"))
%>
  <tr>
    <td class=ctd><%=icounter%></td>
    <td class=ctd <%if Rs("xldh")<>"" then%>alt="修理单号:<%=rs("xldh")%><br><%=Task_Ts(rs("rwlr"))%>"<%else%>alt="任务内容:<br><%=Task_Ts(rs("rwlr"))%>"<%end if%>><%=rs("rwlx")%></td>
    <td class=ctd><%=xjDate(rs("jssj"),1)%></td>
    <td class=ctd><%=rs("zf")%></td>
    <td class=ctd><%=rs("zrr")%></td>
    <%if chkable(3) then%>
    <td class=ctd width=60><a href="ftask_change.asp?id=<%=rs("id")%>" target="_blank">更改</a></td>
    <form action="ftask_indb.asp?action=delete" method=post onsubmit="return confirm('删除ID号为 <%=icounter%> 的零星任务！ 删除后将不能恢复！\n确认吗？');">
      <td class=ctd width=60><input type="submit" value="删除"></td>
      <input type="hidden" name=id value="<%=rs("id")%>">
    </form>
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
    <td align=left> 符合条件共 <%=rs.recordcount%> 个&nbsp;&nbsp;
      每页 <%=rs.pagesize%> 个&nbsp;&nbsp;
      共 <%=Rs.PageCount%> 页&nbsp;&nbsp;
      当前为第 <%=absPageNum%> 页 </td>
    <td align=right> 【
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
