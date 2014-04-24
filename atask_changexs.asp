<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble("3,4")
If Session("userGroup") <> 5 Then Call JsAlert("请联系第五组长进行此辅助任务分配!","atask.asp")
CurPage="调试任务 → 修改调试任务分值系数"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="atask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, s_lsh
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder
s_lsh=request("s_lsh")

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<table class="xtable" cellspacing="0" cellpadding="0" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd" style="border-right-style:none"><%Call SearchLsh()%></td>
    <td class="rtd" style="border-left-style:none"><%Call Taxis()%></td>
  </tr>
  <tr>
    <td class="ctd" height="300" colspan="2"><%If s_lsh="" Then
    		Call atask_nofinished()
    	Else
    		Call ataskAssign()
    	End If%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </td>
  </tr>
</table>
<%
End Sub

Function Taxis()
Dim strO
strO="1"
%>
排序:
<select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;order=&quot; + this.value);'>
  <option value="jssj" selected="selected">完成时间</option>
  <option value="zrr" <%If strOrder="zrr" Then%>selected<%End If%>>责任人</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>流水号</option>
</select>
<%
End Function

Function ataskAssign()
	Dim strXs
	i=1
	If s_lsh="" Then Call TbTopic("请输入修改系数的流水号!") : Exit Function
	strSql="select * from [mantime] where (rwlr='全套调试合格' or rwlr='模头调试合格' or rwlr='定型调试合格' or rwlr like '%精调%' or rwlr like '%验收%'  or rwlr like '%初调%') and lsh='"&s_lsh&"'"
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.open strSql,Conn,1,3

Call TbTopic("修改流水号 " & s_lsh & " 调试任务分值系数")%>
<table width="95%" cellspacing="0" cellpadding="2" class="xtable">
  <tr>
    <th class="th">id</th>
    <th class="th">流水号</th>
    <th class="th">任务内容</th>
    <th class="th" width="*">责任人</th>
    <th class="th" width="*">任务分值</th>
    <th class="th" width="*">实际分值</th>
    <th class="th" width="*">系数</th>
    <th class="th" width="*">完成时间</th>
    <th class="th" width="*">最后修改时间</th>
  </tr>
  <%do while not Rs.Eof %>
  <tr>
    <td class="ctd"><%=i %></td>
    <td class="ctd"><%=rs("lsh")%></td>
    <td class="ctd"><%=rs("rwlr") %></td>
    <td class="ctd"><%=rs("zrr") %></td>
    <td class="ctd"><%=rs("rwfz") %>&nbsp;</td>
    <td class="ctd"><%=rs("fz") %></td>
    <td class="ctd"><%=rs("jc") %></td>
    <td class="ctd"><%=xjDate(rs("jssj"),1)%></td>
    <td class="ctd" alt="<%=Task_Ts(rs("bz"))%>"><%=xjDate(rs("xgsj"),1)%>&nbsp;</td>
  </tr>
  <%
  strXs=rs("jc")
  rs.movenext
  i = i + 1
  loop
%>
</table>
<table width="95%" cellspacing="0" cellpadding="2">
  <form id=frm_cha name=frm_cha action=atask_xsindb.asp method=post onSubmit='return tscheckinf();'>
    <tr >
      <td height="30" align="center"><input type="hidden" name="yxs" value="<%=strXs%>">
        <input type="hidden" name="ylsh" value="<%=s_lsh%>"></td>
    </tr>
    <tr >
      <td align="center">请输入新的系数:
        <input type="text" id="newxs" name="newxs" value="<%=strXs%>" onkeypress="javascript:validationNumber(this, 'u_float', 10, txtFzMsg);" />
        <input type="submit" value=" · 修 改 · " />
        <SPAN id="txtFzMsg"></td>
    </tr>
  </form>
</table>
<%
	rs.close
End Function

Function atask_nofinished()			'具有修改权限的已完成的调试任务
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter,TotalCount, sqlorder, dtend
	absPageNum = 0
	RecordPerPage = 20
	iCounter = 1
	sqlorder = " order by jssj desc"
	If LCase(strOrder) = "zrr" Then sqlorder = " order by zrr desc, lsh desc"
	If LCase(strOrder) = "lsh" Then sqlorder = " order by lsh desc"

	dtend=(dateadd("m",-1,now()))
	dtend=xjDate(year(dtend)&"年"&month(dtend)&"月1日",1)
	strSql="select * from [mantime] where (rwlr='全套调试合格' or rwlr='模头调试合格' or rwlr='定型调试合格' or rwlr like '%精调%' or rwlr like '%验收%'  or rwlr like '%初调%')  and datediff('m',jssj,'"&dtend&"')<=0" & sqlorder
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,1,3
	If (Rs.Eof Or Rs.Bof) Then
		Call TbTopic(dtend&"以后没有可以修改的调试分值系数!") : Exit Function
	End if
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
	i=1
%>
<%Call TbTopic("您共有 "& rs.recordcount &" 个"&dtend&"以后分配的调试分值可以修改系数")%>
<table width="95%" cellspacing="0" cellpadding="2" class="xtable">
  <tr>
    <th class="th">id</th>
    <th class="th">流水号</th>
    <th class="th">任务内容</th>
    <th class="th" width="*">责任人</th>
    <th class="th" width="*">任务分值</th>
    <th class="th" width="*">实际分值</th>
    <th class="th" width="*">系数</th>
    <th class="th" width="*">完成时间</th>
    <th class="th" width="*">修改时间</th>
  </tr>
  <% for absrecordnum = 1 to recordperpage %>
  <tr>
    <td class="ctd"><%=i %></td>
    <td class="ctd"><a href="atask_changexs.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("rwlr") %></td>
    <td class="ctd"><%=rs("zrr") %></td>
    <td class="ctd"><%=rs("rwfz") %>&nbsp;</td>
    <td class="ctd"><%=rs("fz") %></td>
    <td class="ctd"><%=rs("jc") %></td>
    <td class="ctd"><%=xjDate(rs("jssj"),1)%></td>
    <td class="ctd" alt="<%=Task_Ts(rs("bz"))%>"><%=xjDate(rs("xgsj"),1)%>&nbsp;</td>
  </tr>
  <%rs.movenext
  if rs.eof then
  	exit for
  end if
  i = i + 1
  next%>
</table>
<table width="95%" cellpadding="2" cellspacing="0" border="0">
  <tr>
    <td align="left"> 符合条件共 <%=rs.recordcount%> 个&nbsp;&nbsp;
      每页 <%=rs.pagesize%> 个&nbsp;&nbsp;
      共 <%=Rs.PageCount%> 页&nbsp;&nbsp;
      当前为第 <%=absPageNum%> 页 </td>
    <td align="right"> 【
      <%
				if absPageNum > 1 then
					response.write("<a href="&Request.ServerVariables("script_name")&"?ipage="&(abspagenum-1)&" alt='上一页'> ←</a>&nbsp;&nbsp;")
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
						response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&i&">" & i & "</a>&nbsp;")
					end if
				next
				if absPageNum < Rs.PageCount then
					response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&(absPageNum+1)&" alt='下一页'> → </a>&nbsp;")
				end if
			%>
      】
      跳转到:
      <select name="ipage" onchange='location.href(&quot;<%=Request.ServerVariables("script_name")%>?ipage=&quot; + this.value+&quot;&quot;);'>
        <%for i=1 to Rs.PageCount%>
        <%if i = absPageNum then%>
        <option value="<%=i%>" selected="selected">第 <%=i%> 页</option>
        <%else%>
        <option value="<%=i%>">第 <%=i%> 页</option>
        <%end if%>
        <%next%>
      </select>
    </td>
  </tr>
</table>
<%
	rs.close
end function

Function Task_Ts(mystr)
    dim strnew
	if instr(mystr,"||")>0 then
		strnew=split(mystr,"||")
    	for i=0 to ubound(strnew)
			Task_Ts=Task_Ts&strnew(i)&"<br>"
		next
	else
		Task_Ts=mystr
	end if
End Function
%>
<script language="javascript">
function tscheckinf()
{
	var objdm=document.frm_cha
	if (objdm.newxs.value==""){alert("系数不能为空!"); objdm.newxs.focus(); return false;}
	return true;
}
</script>
