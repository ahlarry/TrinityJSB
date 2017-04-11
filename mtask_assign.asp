<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'16:10 2007-1-26-星期五
Call ChkPageAble("3,4")
CurPage="设计任务 → 分配任务书"
strPage="mtask"
Call FileInc(0, "js/mtask.js")
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, sjcount, dbcount
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder
sjcount=0
dbcount=0

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
    <td class="ctd" height="300" colspan="2"><%Call mtaskAssign()%>
      <%Response.Write(XjLine(5,"100%",""))%>
      <%Response.Write(XjLine(1,"100%","#00659c"))%>
      <%Response.Write(XjLine(5,"100%",""))%>
      <%Call task_assign_nofinished()%>
      <%Call jsdb_nofinished()%>
<table width="99%" cellpadding="2" cellspacing="0" border="0">
  <tr>
    <td class="td_ltd">符合条件设计任务书&nbsp;<%=sjcount%>&nbsp;个,技术代表任务书&nbsp;<%=dbcount%>&nbsp;个</td>
  </tr>
</table>
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
  <option value="jhjssj" selected="selected">计划整体完成时间</option>
  <option value="ddh" <%If strOrder="ddh" Then%>selected<%End If%>>订单号</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>流水号</option>
</select>
<%
End Function
Function mtaskAssign()
	Dim s_lsh,s_hth
	s_lsh="" : s_hth=""
	If Trim(request("s_lsh"))<>"" Then s_lsh=Trim(request("s_lsh"))
	If Trim(request("s_hth"))<>"" Then s_hth=Trim(request("s_hth"))
	If s_lsh="" and s_hth="" Then Call TbTopic("请输入或选择待分配任务书的流水号!") : Exit Function
	if s_lsh<>"" Then
		strSql="select * from [mtask] where lsh='"&s_lsh&"'"
		Set Rs=xjweb.Exec(strSql,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("流水号 " & s_lsh & " 任务书不存在! 请重新输入!","mtask_assign.asp")
		Else
			If Not isnull(rs("sjjssj")) then
				Call JsAlert("本任务已经完成,不需要再进行分配!","mtask_assign.asp")
				'17:18 2007-3-20-星期二
			elseif rs("zz")=session("userName") or rs("jgzz")=session("userName") or rs("sjzz")=session("userName") or rs("group")=session("userGroup") or chkable(3) then
				call mtask_assign(rs)
			elseif rs("zz")<>"" Then
				Call JsAlert("本任务的组长是 " & rs("zz") &" ! 请联系 "&rs("zz")&" 进行任务分配!","")
				else
					Call JsAlert("本任务的组长是 " & rs("jgzz") &"(结构)、"& rs("sjzz") &"(设计) ! 请联系组长进行任务分配!","")
			end if
		end if
	else
		strSql="select * from [jsdb] where hth='"&s_hth&"'"
		Set Rs=xjweb.Exec(strSql,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("合同号 " & s_hth & " 的技术代表任务书不存在! 请重新输入!","mtask_assign.asp")
		Else
			If Not isnull(rs("shjssj")) then
				Call JsAlert("本任务已经完成,不需要再进行分配!","mtask_assign.asp")
			elseif rs("zz")=session("userName") or chkable(3) then
					call jsdb_assign(rs)
				else
					Call JsAlert("本任务的组长是 " & rs("zz") &" ! 请联系 "&rs("zz")&" 进行任务分配!","")
			end if
		end if
	end if
	rs.close
end function

function mtask_assign(rs)
	call mtask_fewinfo(rs)
	Response.Write(XjLine(4,"100%",""))
	call mtask_userinfo(rs)
	Response.Write(XjLine(4,"100%",""))
	Response.Write(XjLine(1,"100%","#00659c"))
	Response.Write(XjLine(4,"100%",""))

	if not isnull(rs("fsjs")) then
		If chkable(3) Then
			call mtask_finish(rs)
		else
			Call JsAlert("本任务已复审结束! 请及时完成相应调试单任务!","mtask_assign.asp")
		End If
	else
		select case rs("mjxx")
			case "全套"
				if not isnull(rs("mtbomjs")) and not isnull(rs("dxbomjs")) then
					call mtask_audit(rs)
				else
					If IsNull(Rs("dxshr")) and IsNull(Rs("mtshr")) and (rs("rwlr")="设计") Then
						call mtask_nassign(rs)
					else
						call mtask_assigntask(rs)
					End If
				end if
			case "模头"
				if not isnull(rs("mtbomjs")) then
					call mtask_audit(rs)
				else
					If IsNull(Rs("mtshr")) and (rs("rwlr")="设计") Then
						call mtask_nassign(rs)
					else
						call mtask_assigntask(rs)
					End If
				end if
			case "定型"
				if not isnull(rs("dxbomjs")) then
					call mtask_audit(rs)
				else
					If IsNull(Rs("dxshr")) and (rs("rwlr")="设计") Then
						call mtask_nassign(rs)
					else
						call mtask_assigntask(rs)
					End If
				end if
		end select
	end if
end function		'mtask_assign()

Function mtask_finish(rs)
	if (rs("mjxx")="全套" and  (isnull(rs("mttsdjs")) or isnull(rs("dxtsdjs")))) or  (rs("mjxx")="模头" and isnull(rs("dxtsdjs"))) or  (rs("mjxx")="定型" and isnull(rs("dxtsdjs"))) then
		Call JsAlert("请提醒组长及时完成相应调试单任务!","mtask_assign.asp")
	end if
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%" align="center">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="if(document.all.zrpsjl.value==''){alert('请填写评审记录!');return false;}">
    <tr>
      <td class="ctd" colspan="2" >模具分值:<b><%=Rs("mjzf")%></b></td>
    </tr>
    <tr>
      <td class="ctd" width="15%">评审记录:</td>
      <td class="ctd"><textarea name="zrpsjl" rows="7" cols="60"><%=Rs("psjl")%></textarea></td>
    </tr>
    <tr>
      <td class="ctd">结束时间</td>
      <td class="ctd"><select id="psd" name="psd">
          <%for i = DateAdd("m", -2, date()) to date()%>
          <%if i = date() then%>
          <option value='<%=i%>' selected="selected"><%=i%></option>
          <%else%>
          <option value='<%=i%>'><%=i%></option>
          <%end if%>
          <%next%>
        </select>
        <input name="fplr" type="submit" value="全套结束" />
      </td>
    </tr>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<%
End Function

function mtask_audit(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%" align="center">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="if(document.all.psjl.value==''){alert('请填写评审记录!');return false;}">
    <tr>
      <td class="ctd"  width="15%">评审记录:</td>
      <td class="ctd"><textarea name="psjl" rows="5" cols="80"><%If datediff("d", now, rs("jhjssj")) > 3 Then%>为约束任务安排合理性,系统不允许提前计划3天以上完成.<%End If%>
</textarea></td>
      <td class="ctd"><input name="fplr" type="submit" value="结束复审" <%If datediff("d", now, rs("jhjssj")) > 3 Then%> disabled="disabled" <%End If%> /></td>
    </tr>
    <tr>
    	<td class="ctd"  width="15%">评审人:</td>
    	<td class="ctd">
         <select name="zrr">
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>' <%if session("userName")=c_allzz(i) then %> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
          <option value=“朱钰”>朱钰</option>
          <option value=“徐小停”>徐小停</option>
        </select>
	</td>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<%
end function
'┌─────────────────────────────────────────────┐
function mtask_nassign(rs)
%>
<table border="0" cellpadding="3" cellspacing="0">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style="font-size:14px;font-weight:bold;">分配流水号 <%=rs("lsh")%> 任务书:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
	select case rs("mjxx") & rs("rwlr")
		case "全套设计"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'仅结构组长能分配结构任务
			if Rs("HGJ")>0 Then
				if isnull(rs("mtjgr")) and isnull(rs("dxjgr")) and isnull(rs("gjjgr")) then sel_opt("开始全套结构")
				elseif isnull(rs("mtjgr")) and isnull(rs("dxjgr")) then sel_opt("开始全套结构")
			End if
			if isnull(rs("mtjgr")) then sel_opt("开始模头结构")
			if isnull(rs("dxjgr")) then sel_opt("开始定型结构")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("开始后共挤结构")
			if Rs("HGJ")>0 Then
				if rs("mtjgr")=rs("dxjgr") and rs("dxjgr")=rs("gjjgr") and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) and isnull(rs("gjjgjs")) then sel_opt("结束全套结构")
				elseif (rs("mtjgr")=rs("dxjgr")) and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) then sel_opt("结束全套结构")
			End if
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("结束模头结构")
			if (not isnull(rs("dxjgr"))) and isnull(rs("dxjgjs")) then sel_opt("结束定型结构")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("结束后共挤结构")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjjgjs"))) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) and isnull(rs("mtjgshr")) and isnull(rs("dxjgshr")) and isnull(rs("gjjgshr")) then sel_opt("开始全套结构确认")
				elseif isnull(rs("mtjgshr")) and isnull(rs("dxjgshr")) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) then sel_opt("开始全套结构确认")
			End if
			if isnull(rs("mtjgshr")) and not isnull(rs("mtjgjs")) then sel_opt("开始模头结构确认")
			if isnull(rs("dxjgshr")) and not isnull(rs("dxjgjs")) then sel_opt("开始定型结构确认")
			if isnull(rs("gjjgshr")) and not isnull(rs("gjjgjs")) then sel_opt("开始后共挤结构确认")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'仅设计组长能分配设计任务
			if Rs("HGJ")>0 Then
				if rs("mtjgshr")=rs("dxjgshr") and rs("dxjgshr")=rs("gjjgshr") and isnull(rs("mtjgshjs")) and isnull(rs("dxjgshjs")) and isnull(rs("gjjgshjs")) then sel_opt("结束全套结构确认")
				elseif (rs("mtjgshr")=rs("dxjgshr")) and isnull(rs("mtjgshjs")) and isnull(rs("dxjgshjs")) then sel_opt("结束全套结构确认")
			End if
			if not isnull(rs("mtjgshr")) and isnull(rs("mtjgshjs")) then sel_opt("结束模头结构确认")
			if not isnull(rs("dxjgshr")) and isnull(rs("dxjgshjs")) then sel_opt("结束定型结构确认")
			if not isnull(rs("gjjgshr")) and isnull(rs("gjjgshjs")) then sel_opt("结束后共挤结构确认")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjjgshjs"))) and (not isnull(rs("mtjgshjs"))) and (not isnull(rs("dxjgshjs"))) and isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and isnull(rs("gjsjr")) then sel_opt("开始全套设计")
				elseif isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and (not isnull(rs("mtjgshjs"))) and (not isnull(rs("dxjgshjs"))) then sel_opt("开始全套设计")
			End if
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgshjs")) then sel_opt("开始模头设计")
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgshjs")) then sel_opt("开始定型设计")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgshjs")) then sel_opt("开始后共挤设计")

			if Rs("HGJ")>0 Then
				if rs("mtsjr")=rs("dxsjr") and rs("dxsjr")=rs("gjsjr") and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) and isnull(rs("gjsjjs")) then sel_opt("结束全套设计")
				elseif (rs("mtsjr")=rs("dxsjr")) and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) then sel_opt("结束全套设计")
			End if
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("结束模头设计")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("结束定型设计")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("结束后共挤设计")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjsjjs"))) and (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtsjshr")) and isnull(rs("dxsjshr")) and isnull(rs("gjsjshr")) then sel_opt("开始全套设计确认")
				elseif (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtsjshr")) and isnull(rs("dxsjshr")) then sel_opt("开始全套设计确认")
			End if
			if isnull(rs("mtsjshr")) and not isnull(rs("mtsjjs")) then sel_opt("开始模头设计确认")
			if isnull(rs("dxsjshr")) and not isnull(rs("dxsjjs")) then sel_opt("开始定型设计确认")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjsjshr")) then sel_opt("开始后共挤设计确认")

			if Rs("HGJ")>0 Then
				if rs("mtsjshr")=rs("dxsjshr") and rs("dxsjshr")=rs("gjsjshr") and isnull(rs("mtsjshjs")) and isnull(rs("dxsjshjs")) and isnull(rs("gjsjshjs")) then sel_opt("结束全套设计确认")
				elseif (rs("mtsjshr")=rs("dxsjshr")) and isnull(rs("mtsjshjs")) and isnull(rs("dxsjshjs")) then sel_opt("结束全套设计确认")
			End if
			if not isnull(rs("mtsjshr")) and isnull(rs("mtsjshjs")) then sel_opt("结束模头设计确认")
			if not isnull(rs("dxsjshr")) and isnull(rs("dxsjshjs")) then sel_opt("结束定型设计确认")
			if (not isnull(rs("gjsjshr"))) and isnull(rs("gjsjshjs")) then sel_opt("结束后共挤设计确认")
		End If

'       	if not(isnull(rs("mtsjshjs"))) and not(isnull(rs("dxsjshjs"))) and not(isnull(rs("gjsjshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("全套工艺设计")
'       	if not(isnull(rs("mtsjshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'       	if not(isnull(rs("dxsjshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'       	if not(isnull(rs("gjsjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("共挤工艺设计")
'       	if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr")))  and not(isnull(rs("gjgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("全套工艺审核")
'       	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")
'       	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")
'       	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("共挤工艺审核")
		if isnull(rs("mtbomr")) and not isnull(rs("mtsjshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxsjshjs")) then sel_opt("开始全套BOM")
		if isnull(rs("mtbomr")) and not isnull(rs("mtsjshjs")) then sel_opt("开始模头BOM")
		if isnull(rs("dxbomr")) and not isnull(rs("dxsjshjs")) then sel_opt("开始定型BOM")

		if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("结束全套BOM")
		if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")
		if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")

		case "模头设计"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'仅结构组长能分配结构任务
			if isnull(rs("mtjgr")) then sel_opt("开始模头结构")
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("结束模头结构")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("开始后共挤结构")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("结束后共挤结构")
			if isnull(rs("mtjgshr")) and not isnull(rs("mtjgjs")) then sel_opt("开始模头结构确认")
			if isnull(rs("gjjgshr")) and not isnull(rs("gjjgjs")) then sel_opt("开始后共挤结构确认")
			if not isnull(rs("gjjgshr")) and isnull(rs("gjjgshjs")) then sel_opt("结束后共挤结构确认")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'仅设计组长能分配设计任务
			if (not isnull(rs("mtjgshr"))) and isnull(rs("mtjgshjs")) then sel_opt("结束模头结构确认")
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgshjs")) then sel_opt("开始模头设计")
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("结束模头设计")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgshjs")) then sel_opt("开始后共挤设计")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("结束后共挤设计")
			if isnull(rs("mtsjshr")) and not isnull(rs("mtsjjs")) then sel_opt("开始模头设计确认")
			if not isnull(rs("mtsjshr")) and isnull(rs("mtsjshjs")) then sel_opt("结束模头设计确认")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjsjshr")) then sel_opt("开始后共挤设计确认")
			if (not isnull(rs("gjsjshr"))) and isnull(rs("gjsjshjs")) then sel_opt("结束后共挤设计确认")
		End If

'          	if not(isnull(rs("mtsjshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'          	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")
			if isnull(rs("mtbomr")) and not isnull(rs("mtsjshjs")) then sel_opt("开始模头BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")

		case "定型设计"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'仅结构组长能分配结构任务
			if isnull(rs("dxjgr")) then sel_opt("开始定型结构")
			if not isnull(rs("dxjgr")) and isnull(rs("dxjgjs")) then sel_opt("结束定型结构")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("开始后共挤结构")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("结束后共挤结构")
			if isnull(rs("dxjgshr")) and not isnull(rs("dxjgjs")) then sel_opt("开始定型结构确认")
			if isnull(rs("gjjgshr")) and not isnull(rs("gjjgjs")) then sel_opt("开始后共挤结构确认")
			if not isnull(rs("gjjgshr")) and isnull(rs("gjjgshjs")) then sel_opt("结束后共挤结构确认")
		End If


		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'仅设计组长能分配设计任务
			if not isnull(rs("dxjgshr")) and isnull(rs("dxjgshjs")) then sel_opt("结束定型结构确认")
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgshjs")) then sel_opt("开始定型设计")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("结束定型设计")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgshjs")) then sel_opt("开始后共挤设计")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("结束后共挤设计")
			if isnull(rs("dxsjshr")) and not isnull(rs("dxsjjs")) then sel_opt("开始定型设计确认")
			if not isnull(rs("dxsjshr")) and isnull(rs("dxsjshjs")) then sel_opt("结束定型设计确认")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjsjshr")) then sel_opt("开始后共挤设计确认")
			if (not isnull(rs("gjsjshr"))) and isnull(rs("gjsjshjs")) then sel_opt("结束后共挤设计确认")
		End If

'          	if not(isnull(rs("dxsjshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'          	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")
			if isnull(rs("dxbomr")) and not isnull(rs("dxsjshjs")) then sel_opt("开始定型BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")

		case else
			response.write(rs("mjxx") & rs("rwlr"))
	end select
%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type="submit" value=" 分配 " /></td>
    </tr>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<script language="JavaScript" type="text/javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='结束')
				objdoc.form_assign.zrr.disabled=true;
			else
				objdoc.form_assign.zrr.disabled=false;
		}
		checkjs();

		function checksubinf(frm)
		{
			if (frm.fplr.value==""){alert("请选择分配内容!"); frm.fplr.focus(); return false;}
			if ((frm.zrr.value=="") && (!frm.zrr.disabled)){alert("请选择责任人!"); frm.zrr.focus(); return false;}
			return true;
		}
	</script>
<%
end function
'└─────────────────────────────────────────────┘
function mtask_assigntask(rs)
%>
<table border="0" cellpadding="3" cellspacing="0">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style="font-size:14px;font-weight:bold;">分配流水号 <%=rs("lsh")%> 任务书:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
	select case rs("mjxx") & rs("rwlr")
		case "全套设计"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'仅结构组长能分配结构任务
			if Rs("HGJ")>0 Then
				if isnull(rs("mtjgr")) and isnull(rs("dxjgr")) and isnull(rs("gjjgr")) then sel_opt("开始全套结构")
				elseif isnull(rs("mtjgr")) and isnull(rs("dxjgr")) then sel_opt("开始全套结构")
			End if
			if isnull(rs("mtjgr")) then sel_opt("开始模头结构")
			if isnull(rs("dxjgr")) then sel_opt("开始定型结构")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("开始后共挤结构")

			if Rs("HGJ")>0 Then
				if rs("mtjgr")=rs("dxjgr") and rs("dxjgr")=rs("gjjgr") and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) and isnull(rs("gjjgjs")) then sel_opt("结束全套结构")
				elseif (rs("mtjgr")=rs("dxjgr")) and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) then sel_opt("结束全套结构")
			End if
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("结束模头结构")
			if (not isnull(rs("dxjgr"))) and isnull(rs("dxjgjs")) then sel_opt("结束定型结构")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("结束后共挤结构")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'仅设计组长能分配设计任务
			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjjgjs"))) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) and isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and isnull(rs("gjsjr")) then sel_opt("开始全套设计")
				elseif isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) then sel_opt("开始全套设计")
			End if
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgjs")) then sel_opt("开始模头设计")
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgjs")) then sel_opt("开始定型设计")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgjs")) then sel_opt("开始后共挤设计")

			if Rs("HGJ")>0 Then
				if rs("mtsjr")=rs("dxsjr") and rs("dxsjr")=rs("gjsjr") and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) and isnull(rs("gjsjjs")) then sel_opt("结束全套设计")
				elseif (rs("mtsjr")=rs("dxsjr")) and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) then sel_opt("结束全套设计")
			End if
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("结束模头设计")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("结束定型设计")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("结束后共挤设计")
		End If

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjsjjs"))) and (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) and isnull(rs("gjshr")) then sel_opt("开始全套审核")
				elseif (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) then sel_opt("开始全套审核")
			End if
			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("开始模头审核")
			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("开始定型审核")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjshr")) then sel_opt("开始后共挤审核")

			if Rs("HGJ")>0 Then
				if rs("mtshr")=rs("dxshr") and rs("dxshr")=rs("gjshr") and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) and isnull(rs("gjshjs")) then sel_opt("结束全套审核")
				elseif (rs("mtshr")=rs("dxshr")) and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) then sel_opt("结束全套审核")
			End if
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("结束模头审核")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("结束定型审核")
			if (not isnull(rs("gjshr"))) and isnull(rs("gjshjs")) then sel_opt("结束后共挤审核")

'     	  	if not(isnull(rs("mtshjs"))) and not(isnull(rs("dxshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("全套工艺设计")
'   	    	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'      	 	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'       		if not(isnull(rs("gjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("共挤工艺设计")
'      	 	if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("全套工艺审核")
'      	 	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")
'     	  	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")
'     	  	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("共挤工艺审核")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始全套BOM")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("开始模头BOM")
			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始定型BOM")
			if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("结束全套BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")

		case "全套复改"
			if Rs("HGJ")>0 Then
				if isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and isnull(rs("gjsjr")) then sel_opt("开始全套复改")
				elseif isnull(rs("mtsjr")) and isnull(rs("dxsjr")) then sel_opt("开始全套复改")
			End if
			if isnull(rs("mtsjr")) then sel_opt("开始模头复改")
			if isnull(rs("dxsjr")) then sel_opt("开始定型复改")
			if Rs("HGJ")>0 and isnull(rs("gjsjr")) then sel_opt("开始后共挤复改")

			if Rs("HGJ")>0 Then
				if rs("mtsjr")=rs("dxsjr") and rs("dxsjr")=rs("gjsjr") and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) and isnull(rs("gjsjjs")) then sel_opt("结束全套复改")
				elseif (rs("mtsjr")=rs("dxsjr")) and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) then sel_opt("结束全套复改")
			End if
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("结束模头复改")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("结束定型复改")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("结束后共挤复改")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjsjjs"))) and (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) and isnull(rs("gjshr")) then sel_opt("开始全套审核")
				elseif (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) then sel_opt("开始全套审核")
			End if
			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("开始模头审核")
			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("开始定型审核")
			if isnull(rs("gjshr")) and not isnull(rs("gjsjjs")) then sel_opt("开始后共挤审核")

			if Rs("HGJ")>0 Then
				if rs("mtshr")=rs("dxshr") and rs("dxshr")=rs("gjshr") and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) and isnull(rs("gjshjs")) then sel_opt("结束全套审核")
				elseif (rs("mtshr")=rs("dxshr")) and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) then sel_opt("结束全套审核")
			End if
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("结束模头审核")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("结束定型审核")
			if not isnull(rs("gjshr")) and isnull(rs("gjshjs")) then sel_opt("结束后共挤审核")

'  	     	if not(isnull(rs("mtshjs"))) and not(isnull(rs("dxshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("全套工艺设计")
'    	   	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'    	   	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'    	   	if not(isnull(rs("gjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("共挤工艺设计")
'   	    	if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("全套工艺审核")
'   	    	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")
'   	    	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")
'   	    	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("共挤工艺审核")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始全套BOM")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("开始模头BOM")
			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始定型BOM")

			if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("结束全套BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")

		case "全套复查"
			if Rs("HGJ")>0 Then
				if isnull(rs("mtshr")) and isnull(rs("dxshr")) and isnull(rs("gjshr")) then sel_opt("开始全套复查")
				elseif isnull(rs("mtshr")) and isnull(rs("dxshr")) then sel_opt("开始全套复查")
			End if
			if isnull(rs("mtshr")) then sel_opt("开始模头复查")
			if isnull(rs("dxshr")) then sel_opt("开始定型复查")
			if Rs("HGJ")>0 and isnull(rs("gjshr")) then sel_opt("开始后共挤复查")

			if Rs("HGJ")>0 Then
				if rs("mtshr")=rs("dxshr") and rs("dxshr")=rs("gjshr") and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) and isnull(rs("gjshjs")) then sel_opt("结束全套复查")
				elseif (rs("mtshr")=rs("dxshr")) and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) then sel_opt("结束全套复查")
			End if
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("结束模头复查")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("结束定型复查")
			if not isnull(rs("gjshr")) and isnull(rs("gjshjs")) then sel_opt("结束后共挤复查")

'    	   	if not(isnull(rs("mtshjs"))) and not(isnull(rs("dxshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("全套工艺设计")
'       		if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'       		if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'       		if not(isnull(rs("gjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("共挤工艺设计")
'       		if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("全套工艺审核")
'       		if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")
'      	 	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")
'      	 	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("共挤工艺审核")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始全套BOM")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("开始模头BOM")
			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始定型BOM")

			if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("结束全套BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")

		case "模头设计"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'仅结构组长能分配结构任务
			if isnull(rs("mtjgr")) then sel_opt("开始模头结构")
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("结束模头结构")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'仅设计组长能分配设计任务
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgjs")) then sel_opt("开始模头设计")
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("结束模头设计")
		End If

			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("开始模头审核")
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("结束模头审核")

'      	 	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'       	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("开始模头BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")

		case "模头复改"
			if isnull(rs("mtsjr")) then sel_opt("开始模头复改")
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("结束模头复改")

			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("开始模头审核")
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("结束模头审核")

' 	      	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'   	    	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("开始模头BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")

		case "模头复查"
			if isnull(rs("mtshr")) then sel_opt("开始模头复查")
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("结束模头复查")

'    	   	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("模头工艺设计")
'   	    	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("模头工艺审核")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("开始模头BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("结束模头BOM")

		case "定型设计"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'仅结构组长能分配结构任务
			if isnull(rs("dxjgr")) then sel_opt("开始定型结构")
			if not isnull(rs("dxjgr")) and isnull(rs("dxjgjs")) then sel_opt("结束定型结构")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'仅设计组长能分配设计任务
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgjs")) then sel_opt("开始定型设计")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("结束定型设计")
		End If

			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("开始定型审核")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("结束定型审核")

'       	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'       	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")

			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始定型BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")

		case "定型复改"
			if isnull(rs("dxsjr")) then sel_opt("开始定型复改")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("结束定型复改")

			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("开始定型审核")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("结束定型审核")

'   	    	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'     	  	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")

			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始定型BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")

		case "定型复查"
			if isnull(rs("dxshr")) then sel_opt("开始定型复查")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("结束定型复查")

'   	    	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("定型工艺设计")
'     	  	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("定型工艺审核")

			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("开始定型BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("结束定型BOM")
		case else
			response.write(rs("mjxx") & rs("rwlr"))
	end select
%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type="submit" value=" 分配 " /></td>
    </tr>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<script language="JavaScript" type="text/javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='结束')
				objdoc.form_assign.zrr.disabled=true;
			else
				objdoc.form_assign.zrr.disabled=false;
		}
		checkjs();

		function checksubinf(frm)
		{
			if (frm.fplr.value==""){alert("请选择分配内容!"); frm.fplr.focus(); return false;}
			if ((frm.zrr.value=="") && (!frm.zrr.disabled)){alert("请选择责任人!"); frm.zrr.focus(); return false;}
			return true;
		}
	</script>
<%
end function

function sel_opt(str)
%>
<option value="<%=str%>"><%=str%></option>
<%
end function

function jsdb_assign(rs)
%>
<table border="0" cellpadding="3" cellspacing="0">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style="font-size:14px;font-weight:bold;">分配合同号 <%=rs("hth")%> 技术代表任务书:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
		if isnull(rs("sjr")) then sel_opt("开始技术代表设计")
		if not isnull(rs("sjr")) and isnull(rs("sjjssj")) then sel_opt("结束技术代表设计")

		if isnull(rs("shr")) and not isnull(rs("sjjssj")) then sel_opt("开始技术代表审核")
		if not isnull(rs("shr")) and isnull(rs("shjssj")) then sel_opt("结束技术代表审核")
%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type="submit" value=" 分配 " /></td>
    </tr>
    <input type="hidden" name="hth" value="<%=rs("hth")%>" />
  </form>
</table>
<script language="JavaScript" type="text/javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='结束')
				objdoc.form_assign.zrr.disabled=true;
			else
				objdoc.form_assign.zrr.disabled=false;
		}
		checkjs();

		function checksubinf(frm)
		{
			if (frm.fplr.value==""){alert("请选择分配内容!"); frm.fplr.focus(); return false;}
			if ((frm.zrr.value=="") && (!frm.zrr.disabled)){alert("请选择责任人!"); frm.zrr.focus(); return false;}
			return true;
		}
	</script>
<%
End function

function task_assign_nofinished()			'具有分配权限的未完成的任务书2013/1/19 10:59
	Dim sqlorder
	sqlorder = " order by jhjssj"
	If LCase(strOrder) = "ddh" Then sqlorder = " order by ddh desc, lsh desc"
	If LCase(strOrder) = "lsh" Then sqlorder = " order by lsh desc"

	If chkable(3) Then
		strSql="select * from [mtask] where (not isnull(fsr)) and isnull(sjjssj) " & sqlorder
	else
		strSql="select * from [mtask] where (zz='"&session("userName")&"' or jgzz='"&session("userName")&"' or sjzz='"&session("userName")&"'or [group]="&session("userGroup")&") and isnull(fsjs) and isnull(sjjssj)" & sqlorder
	End If
	call xjweb.exec(strSql,0)
	rs.open strsql,conn,1,3
	if (rs.eof or rs.bof) then Call TbTopic("暂时没有待分配设计任务书!") : Exit Function
	dim icounter
	Call CutLine()		'显示图例
	icounter = 1
	sjcount=Rs.recordcount
%>
<%Call TbTopic("您共有 " & sjcount & " 个待分配的设计任务书流程")%>
<table width="100%" cellspacing="0" class="xtable">
  <tr>
    <td class="th" width="*">ID</td>
    <td class="th" width="*">流水号</td>
    <td class="th" width="*">单位名</td>
    <td class="th" width="7%">断面名</td>
    <td class="th" width="7%">&nbsp;&nbsp;组&nbsp;&nbsp;&nbsp;&nbsp;长&nbsp;&nbsp;</td>
    <td class="th" width="11%">计划结构完成</td>
    <td class="th" width="11%">计划整体完成</td>
    <td class="th" width="*">任务</td>
    <td class="th" width="7%">结构</td>
    <td class="th" width="7%">设计</td>
    <td class="th" width="7%">审核</td>
    <td class="th" width="*">工艺</td>
    <td class="th" width="7%">BOM</td>
  </tr>
  <%do while not rs.eof
	'结构\设计独立计时的图例15:54 2007-4-1-星期日
	'最后一天不算延期，实际要求设计周期应该为datediff("d", 计划开始, 计划结束)+1
	'根据考核办法结构周期=(设计周期-1)/2，因此结构周期＝设计周期/2
  Dim jgjgsj, sjjgsj, Tmpsj, ijgsj, isj
If rs("jhkssj")<>"" Then
	Tmpsj=rs("jhkssj")
else
	Tmpsj=rs("rwxdsj")
End If
  jgjgsj=datediff("d", Tmpsj, rs("jhjssj"))/2
  sjjgsj=datediff("d", Tmpsj, rs("jhjssj"))+1-jgjgsj
If IsNull(rs("jhjgsj")) Then
	isj=INT(datediff("d", rs("jhkssj"), rs("jhjssj"))/2)
	ijgsj=dateadd("d",isj,rs("jhkssj"))
else
	isj=rs("jhjssj")
	ijgsj=rs("jhjgsj")
End if
  %>
  <tr>
    <td class="ctd"><%=icounter%></td>
    <td class="ctd"><a href="mtask_assign.asp?s_LSH=<%=rs("LSH")%>"><%=rs("LSH")%></a></td>
    <td class="ctd"><%=rs("DWMC")%></td>
    <td class="ctd"><%=rs("DMMC")%></td>
    <td class="ctd">
    <%If Rs("zz")<>"" Then
    		Response.Write(rs("zz"))
   	ElseIf rs("jgzz")=rs("sjzz") Then
    		Response.Write(rs("jgzz"))
    else
    	Response.Write(rs("jgzz")&"<br>"&rs("sjzz"))
     End If%>
    </td>
    <%if isNull(rs("JHKSSJ")) then%>
    <td class="ctd">&nbsp;/&nbsp;</td>
    <%else%>
    <td class="ctd" alt='计划开始日期:<%=rs("jhkssj")%>'><%=rs("jhjgsj")%>&nbsp;</td>
    <%end if%>
    <%if isNull(rs("SJJSSJ")) then%>
    <td class="ctd" alt='尚未完成'><%= rs("JHJSSJ") %></td>
    <%else%>
    <td class="ctd" alt='实际结束日期:<%=rs("SJJSSJ")%>'><%= rs("JHJSSJ") %></td>
    <%end if%>
    <td class="ctd"><%= rs("MJXX") & rs("RWLR") %></td>
    <%select case rs("mjxx") & rs("rwlr")%>
    <%case "全套设计"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%=jgjgsj%>
      <%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <% If not(isnull(rs("gjjgks"))) Then call DisTdjg(rs("gjjgks"),rs("gjjgjs"),ijgsj,rs)%>
    </td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%=sjjgsj%>
      <%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <% If not(isnull(rs("gjsjks"))) Then call DisTdjg(rs("gjsjks"),rs("gjsjjs"),isj,rs)%>
    </td>
    <td class="ctd"><%If not(isnull(rs("mtshr"))) or not(isnull(rs("dxshr"))) Then
      		call distd(rs("mtshks"),rs("mtshjs"),0,rs)
      		call distd(rs("dxshks"),rs("dxshjs"),0,rs)
     		If not(isnull(rs("gjshr"))) Then call distd(rs("gjshks"),rs("gjshjs"),0,rs) End If
     	else
     		call DisTdjg(rs("mtjgshks"),rs("mtjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("dxjgshks"),rs("dxjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("mtsjshks"),rs("mtsjshjs"),isj,rs)
       		call DisTdjg(rs("dxsjshks"),rs("dxsjshjs"),isj,rs)
       		If not(isnull(rs("gjjgr"))) Then call DisTdjg(rs("gjjgshks"),rs("gjjgshjs"),ijgsj,rs) End If
       		If not(isnull(rs("gjsjr"))) Then call DisTdjg(rs("gjsjshks"),rs("gjsjshjs"),isj,rs) End If
      End If%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "全套复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
      <%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "全套复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "模头设计"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%=jgjgsj%> </td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%=sjjgsj%> </td>
    <td class="ctd"><%If not(isnull(rs("mtshr"))) Then
      		call distd(rs("mtshks"),rs("mtshjs"),0,rs)
     	else
     		call DisTdjg(rs("mtjgshks"),rs("mtjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("mtsjshks"),rs("mtsjshjs"),isj,rs)
      End If%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
    </td>
    <%case "模头复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
    </td>
    <%case "模头复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
    </td>
    <%case "定型设计"%>
    <td class="ctd"><%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <%=jgjgsj%> </td>
    <td class="ctd"><%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <%=sjjgsj%> </td>
    <td class="ctd"><%If not(isnull(rs("dxshr"))) Then
      		call distd(rs("dxshks"),rs("dxshjs"),0,rs)
     	else
       		call DisTdjg(rs("dxjgshks"),rs("dxjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("dxsjshks"),rs("dxsjshjs"),isj,rs)
      End If%>
    </td>
    <td class="ctd">
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "定型复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd">
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "定型复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd">
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%end select%>
  </tr>
  <%
		rs.movenext
		icounter = icounter + 1
	loop
	rs.close
%>
</table>
<%End function

function jsdb_nofinished()			'具有分配权限的未完成的技术代表任务书
	Dim sqlorder, icounter
	If LCase(strOrder) = "ddh" or LCase(strOrder) = "lsh" Then
		sqlorder = " order by hth desc"
	else
		sqlorder = " order by jhjssj"
	End If

	strSql="select * from [jsdb] where zz='"&session("userName")&"' and isnull(shjssj)" & sqlorder
'	Set Rs=xjweb.Exec(strSql,1)
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.open strSql,Conn,1,3
	if (Rs.eof or Rs.bof) then Call TbTopic("暂时没有待分配技术代表任务书!") : Exit Function
	icounter = 1
	dbcount=Rs.recordcount
	Call TbTopic("您共有 " & dbcount & " 个待分配的技术代表任务书流程")
%>
<table width="100%" cellspacing="0" class="xtable">
  <tr>
    <td class="th" width="*">ID</td>
    <td class="th" width="*">订单号</td>
    <td class="th" width="15%">客户名称</td>
    <td class="th" width="*">任务内容</td>
    <td class="th" width="7%">&nbsp;&nbsp;组&nbsp;&nbsp;&nbsp;&nbsp;长&nbsp;&nbsp;</td>
    <td class="th" width="11%">计划完成时间</td>
    <td class="th" width="7%">设计</td>
    <td class="th" width="7%">审核</td>
  </tr>
  <%do while not rs.eof%>
  <tr>
      <td class="ctd"><%=icounter%></td>
    <td class="ctd"><a href="mtask_assign.asp?s_hth=<%=rs("hth")%>"><%=rs("hth")%></a></td>
    <td class="ctd"><%=rs("khmc")%></td>
    <td class="ctd"><%=rs("rwnr")%></td>
    <td class="ctd"><%=rs("zz")%></td>
    <td class="ctd"><%=rs("jhjssj")%></td>
    <td class="ctd"><%call distd(rs("sjkssj"),rs("sjjssj"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("shkssj"),rs("shjssj"),0,rs)%></td>
</tr>
  <%
		rs.movenext
		icounter = icounter + 1
	loop
	rs.close
%>
</table>
<%end function%>