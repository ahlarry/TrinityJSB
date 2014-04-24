<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'14:32 2007-1-22-星期一
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="设计任务 → 技术代表"
strPage="mtask"
xjweb.header()
Call TopTable()

Dim strFeedBack, strHth, strKhmc, strzuz, strOrder, strTerm, striPage
strHth=Trim(Request("hth"))
strKhmc=Trim(Request("khmc"))
strzuz=Trim(Request("zuz"))
strOrder=Trim(Request("order"))
strTerm=Trim(Request("term"))
striPage =request("ipage")

	strFeedBack=""
	If strHth<>"" Then strFeedBack="&hth="&strHth
	If strKhmc<>"" Then strFeedBack="&khmc="&strKhmc&strFeedBack
	If strzuz<>"" Then strFeedBack="&zuz="&strzuz&strFeedBack
	If strOrder<>"" Then strFeedBack="&order="&strOrder&strFeedBack
	If strTerm<>"" Then strFeedBack="&term="&strTerm&strFeedBack

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()

%>
<table class="xtable" cellspacing="0" cellpadding="2" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd"><%Call SearchInfo()%></td>
  </tr>
  <tr>
    <td class="ctd" height="280"><%Call TaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%></td>
  </tr>
</table>
<%
End Sub

Function SearchInfo()
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td>&nbsp;&nbsp;合同号:
        <input type="text" name="hth" value="<%=strHth%>" size="15" /></td>
      <td>客户名称:
        <input type="text" name="khmc" value="<%=strKhmc%>" size="8" /></td>
      <td>组长:
        <input type="text" name="zuz" value="<%=strzuz%>" size="8" /></td>
      <td align="center"><input type="submit" value=" · 查 找 · " /></td>
      <td>排序:
        <select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;hth=<%= strhth%>&amp;khmc=<%=strKhmc%>&amp;term=<%=strterm%>&amp;order=&quot; + this.value);'>
          <%If strOrder = "hth" Then%>
          <option value="jhjssj">计划完成时间</option>
          <option value="hth" selected="selected">合同号</option>
          <option value="zz">组长</option>
          <%ElseIf strOrder="zz" Then%>
          <option value="jhjssj">计划完成时间</option>
          <option value="hth">合同号</option>
          <option value="zz" selected="selected">组长</option>
          <%Else%>
          <option value="jhjssj" selected="selected">计划完成时间</option>
          <option value="hth">合同号</option>
          <option value="zz">组长</option>
          <%End If%>
        </select></td>
      <td>条件:
        <select name="term" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;hth=<%= strhth%>&amp;khmc=<%=strkhmc%>&amp;order=<%=strorder%>&amp;term=&quot; + this.value);'>
          <%select case strterm%>
          <%case "no"%>
          <option value="no" selected="selected">未完成</option>
          <option value="ok">已完成</option>
          <option value="all">全部</option>
          <%case "ok"%>
          <option value="no">未完成</option>
          <option value="ok" selected="selected">已完成</option>
          <option value="all">全部</option>
          <%case else%>
          <option value="no">未完成</option>
          <option value="ok">已完成</option>
          <option value="all" selected="selected">全部</option>
          <%end select%>
        </select></td>
    </tr>
  </form>
</table>
<%
End Function

Function TaskList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter
	absPageNum = 0
	strSql = ""
	RecordPerPage=20 		'每页显示记录数
	If strHth <> "" Then
			strSql  = " hth like '%"&strHth&"%' "
	End If
	If strkhmc <> "" Then
		If strSql <> "" Then
			strSql = " khmc like '%"&strkhmc&"%' and " & strSql
		Else
			strSql  = " khmc like '%"&strkhmc&"%' "
		End If
	End If
	If strzuz <> "" Then
		If strSql <> "" Then
			strSql = " zz like '%"&strzuz&"%'  and " & strSql
		Else
			strSql  = " zz like '%"&strzuz&"%' "
		End If
	End If

	Select Case strTerm
			Case "no"
				If strSql <> "" Then
					strSql = " isnull(shjssj) and " & strSql
				Else
					strSql  = " isnull(shjssj) "
				End If
			Case "ok"
				If strSql <> "" Then
					strSql = " not(isnull(shjssj)) and " & strSql
				Else
					strSql  = " not(isnull(shjssj)) "
				End If
			Case Else
	End Select

	If Request("strSql") <> "" Then
		strSql = request("strSql")
	End If

	Dim sqlOrder
	sqlOrder = " order by jhjssj desc, hth desc"
	If LCase(strOrder) = "jhjssj" Then sqlOrder = " order by jhjssj desc, hth desc"
	If LCase(strOrder) = "hth" Then sqlOrder = " order by hth desc,zz desc"
	If LCase(strOrder) = "zz" Then sqlOrder = " order by zz desc, hth desc, jhjssj desc"
	If strSql <> "" Then
		strSql = "select * from [jsdb] where "& strSql & sqlOrder
	Else
		strSql = "select * from [jsdb] "& sqlOrder
	End If
	Set Rs = Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Call xjweb.Exec("",-1)
	Rs.open strSql,Conn,3,3
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("指定条件的任务书不存在,请放宽筛选条件!","")
	End If
	Rs.PageSize = RecordPerPage

	If Trim(Request("iPage")) <> ""  Then
		If IsNumeric(Trim(Request("iPage"))) Then
			If Trim(Request("iPage")) <= 0 Then
				absPageNum=1
			ElseIf CLng(Trim(Request("iPage"))) > Rs.Pagecount Then
				absPageNum = Rs.Pagecount
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

	If absPageNum>Rs.PageCount Then absPageNum=Rs.PageCount
	Rs.absolutePage = absPageNum
	Call CutLine()		'显示图例
	Call TbTopic("挤出模具厂技术代表任务列表")
	iCounter = (absPageNum - 1) * RecordPerPage + 1
%>
<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable" align="center">
  <tr>
    <th class="th" width="25">id </th>
    <th class="th" width="60">合同号 </th>
    <th class="th" width="80">客户名称 </th>
    <th class="th" width="*">任务内容 </th>
    <th class="th" width="40">组长 </th>
    <th class="th" width="80">计划完成 </th>
    <th class="th" width="*">设计 </th>
    <th class="th" width="*">审核 </th>
  </tr>
  <%
  	For absRecordNum = 1 To RecordPerPage		'分页
%>
  <tr>
    <td class="ctd"><%=iCounter%></td>
    <td class="ctd"><%=rs("hth")%></a></td>
    <td class="ctd" ><%=rs("khmc")%></a></td>
    <td class="ctd" alt="<%=rs("rwnr")%>"><%=xjweb.StringCut(rs("rwnr"),16)%></td>
    <td class="ctd"><%=rs("zz")%></td>
    <td class="ctd"><%=rs("jhjssj")%></td>
    <td class="ctd"><%call distd(rs("sjkssj"),rs("sjjssj"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("shkssj"),rs("shjssj"),0,rs)%></td>
    <%rs.movenext%>
    <%if rs.eof then
  		exit for
  	end if
  iCounter=iCounter+1
  next%>
</table>
<table width="98%" cellpadding="2" cellspacing="0" border="0" align="center">
  <tr>
    <td align="left"> 符合条件共 <%=Rs.RecordCount%> 个&nbsp;&nbsp;
      每页 <%=Rs.PageSize%> 个&nbsp;&nbsp;
      共 <%=Rs.PageCount%> 页&nbsp;&nbsp;
      当前为第 <%=absPageNum%> 页 </td>
    <td align="right"> 【
      <%
				If absPageNum>1 Then
					Response.Write("<a href="&request.servervariables("script_name")&"?ipage="&(abspagenum-1)&strFeedBack&" title='上一页'> ←</a>&nbsp;&nbsp;")
				End If
				Dim iStart,iEnd
				If absPageNum<4 Then
					iStart=1
				Else
					iStart=absPageNum-3
				End If
				If absPageNum<Rs.PageCount-3 Then
					iEnd = absPageNum + 3
				Else
					iEnd = Rs.PageCount
				End If
				For i=iStart to iEnd
					If i=absPageNum Then
						Response.Write("&nbsp;<font style=font-size:11pt;><b>"&  i & "</b></font>&nbsp;")
					Else
						Response.Write("&nbsp;<a href="&Request.ServerVariables("SCRIPT_NAME")&"?iPage="&i&strFeedBack&">" & i & "</a>&nbsp;")
					End If
				Next
				If absPageNum<Rs.PageCount Then
					Response.Write("&nbsp;<a href="&Request.ServerVariables("SCRIPT_NAME")&"?iPage="&(abspagenum+1)&strFeedBack&" title='下一页'> → </a>&nbsp;")
				End If
			%>
      】
      跳转到:
      <select name="ipage" onchange='location.href(&quot;<%=Request.ServerVariables("SCRIPT_NAME")%>?ipage=&quot; + this.value +&quot;<%=strFeedback%>&quot;);'>
        <%For i=1 To Rs.PageCount%>
        <%If i=absPageNum Then%>
        <option value="<%=i%>" selected="selected">第 <%=i%> 页</option>
        <%Else%>
        <option value="<%=i%>">第 <%=i%> 页</option>
        <%End If%>
        <%Next%>
      </select></td>
  </tr>
</table>
<%
End Function
%>
