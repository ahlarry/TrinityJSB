<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
Call ChkPageAble("3,4,6")
CurPage="调试任务 → 分配调试任务"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="atask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <tr>
    <td class="ctd" style="border-right-style:none"><%Call SearchLsh()%></td>
    <td class="rtd" style="border-left-style:none"><%Call Taxis()%></td>
  </tr>
  <Tr>
    <Td class=ctd height=300 colspan="2"><%Call ataskAssign()%>
      <%Response.Write(XjLine(10,"100%",""))%>
      <%Response.Write(XjLine(1,"100%",web_info(12)))%>
      <%Call atask_nofinished() %>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function Taxis()
Dim strO
strO="1"
%>
排序:
<select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;order=&quot; + this.value);'>
  <option value="jhjssj" selected="selected">调试单计划完成时间</option>
  <option value="ddh" <%If strOrder="ddh" Then%>selected<%End If%>>订单号</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>流水号</option>
</select>
<%
End Function

'调试任务分两组,一组为各组长分配(做调试单),另一种为第五组长分配(调试及调试信息整理)
Function ataskAssign()
	Dim s_lsh
	s_lsh=""
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("请输入分配辅助任务的流水号!") : Exit Function

	'strSql="select a.*, b.*, a.lsh as lsh from [mtask] a, [mtask] b where a.lsh='"&s_lsh&"' and a.lsh=b.lsh and (not isnull(sjjssj)) and not mjjs"
	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书不存在!","atask_assign.asp") : Exit Function
	ElseIf IsNull(Rs("fsjs")) Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书正在设计中!","atask_assign.asp") : Exit Function
	ElseIf Rs("mjjs") Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书已经全部完成!","atask_assign.asp") : Exit Function
	Else
		Select Case Rs("mjxx")
			Case "全套"
				If not(isnull(rs("mttsdjs"))) and not(isnull(rs("dxtsdjs")))	then
					call group5_assign(rs)
				else
					'response.write rs("group")
					call group_assign(rs)
				end if
			Case "模头"
				if not(isnull(rs("mttsdjs"))) then
					call group5_assign(rs)
				else
					call group_assign(rs)
				end if
			Case "定型"
				if not(isnull(rs("dxtsdjs"))) then
					call group5_assign(rs)
				else
					call group_assign(rs)
				end if
		End Select
	End If
	Rs.Close
End Function

Function mtask_info(rs)
	Call mtask_fewinfo(rs)
	Response.Write(xjLine(4, "100%", ""))%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%" align="center">
  <tr>
    <td class="rtd" width="15%">厂内调试</td>
    <td colspan="3" class="ltd" width="35%"><%if rs("cnts") then%>
      是
      <%else%>
      &nbsp;/
      <%end if%></td>
    <td class="rtd" width="15%">调试类别</td>
    <% If Rs("cnts") Then%>
    <%If Not(isnull(Rs("tslb"))) Then%>
    <td class="ltd"><%=Rs("tslb")%></td>
    <%Else%>
    <td class="ltd">&nbsp;/</td>
    <%End If%>
    <%Else%>
    <%If Rs("beit") Then%>
    <td class="ltd">北调</td>
    <%Else%>
    <td class="ltd">&nbsp;/</td>
    <%End If%>
    <%End If%>
  </tr>
</table>
	<%Response.Write(xjLine(4, "100%", ""))
	Call atask_userinfo(rs)
	Response.Write(xjLine(8, "100%", ""))
	Response.Write(xjLine(1, "100%", "#00659c"))
	Response.Write(xjLine(8, "100%", ""))
End Function

Function group_assign(rs)
	If rs("group") = Session("userGroup") Or Rs("zz")=Session("UserName") Or Rs("jgzz")=Session("UserName") Or Rs("sjzz")=Session("UserName") Then
		Call mtask_info(Rs)
	%>
<table border=0 cellpadding=3 cellspacing=0 align="center">
  <form name="form_assign" action="atask_assignindb.asp" method=post onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style=font-size:14px;font-weight:bold;>分配流水号 <%=rs("lsh")%> 辅助任务:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
				select case rs("mjxx")
					case "全套"
						if isnull(rs("mttsdr")) then sel_opt("开始模头调试单")
						if isnull(rs("dxtsdr")) then sel_opt("开始定型调试单")
						if isnull(rs("mttsdr")) and isnull(rs("dxtsdr")) then sel_opt("开始全套调试单")

						if (not isnull(rs("mttsdr"))) and isnull(rs("mttsdjs")) then sel_opt("结束模头调试单")
						if not isnull(rs("dxtsdr")) and isnull(rs("dxtsdjs")) then sel_opt("结束定型调试单")
						if (rs("mttsdr")=rs("dxtsdr")) and isnull(rs("mttsdjs")) and isnull(rs("dxtsdjs")) then sel_opt("结束全套调试单")

					case "模头"
						if isnull(rs("mttsdr")) then sel_opt("开始模头调试单")
						if (not isnull(rs("mttsdr"))) and isnull(rs("mttsdjs")) then sel_opt("结束模头调试单")

					case "定型"
						if isnull(rs("dxtsdr")) then sel_opt("开始定型调试单")
						if not isnull(rs("dxtsdr")) and isnull(rs("dxtsdjs")) then sel_opt("结束定型调试单")
					case else
						response.write(rs("mjxx"))
				end select
				%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
          <option value="TT调试员">TT调试员</option>
          <option value="TB调试员">TB调试员</option>
        </select>
      </td>
      <td><input type=submit value=" 分配 "></td>
    </tr>
    <input type=hidden name=lsh value=<%=rs("lsh")%>>
  </form>
</table>
<%
		call atask_js()
	Else
		Call JsAlert("本任务的组长是 " & rs("zz") & rs("jgzz") &"(结构)、"& rs("sjzz") &"(设计) ! 请联系组长进行任务分配!","")
	End If
end function

Function group5_assign(rs)
	If Session("userGroup") <> 5 Then
		Call JsAlert("请联系第五组长进行此辅助任务分配!","atask_assign.asp")
	Else
		Call mtask_info(Rs)
	%>
<table border=0 cellpadding=3 cellspacing=0 align="center">
  <form name="form_assign" action="atask_assignindb.asp" method=post onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style=font-size:14px;font-weight:bold;>分配流水号 <%=rs("lsh")%> 辅助任务:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
					select case rs("mjxx")
						case "全套"
							if isnull(rs("mttsr")) and not isnull(rs("mttsdjs")) and isnull(rs("dxtsr")) and not isnull(rs("dxtsdjs")) then sel_opt("开始全套调试")
							if isnull(rs("mttsr")) and not isnull(rs("mttsdjs")) then sel_opt("开始模头调试")
							if isnull(rs("dxtsr")) and not isnull(rs("dxtsdjs")) then sel_opt("开始定型调试")

							if (rs("mttsr")=rs("dxtsr")) and isnull(rs("mttsjs"))  and isnull(rs("dxtsjs")) then
								sel_opt("结束全套调试")
								sel_opt("全套厂内初调")
								sel_opt("全套厂外精调")
								sel_opt("全套预验收或寄样")
								sel_opt("全套来厂验收")
							End If
							if not isnull(rs("mttsr")) and isnull(rs("mttsjs")) then
								sel_opt("结束模头调试")
								sel_opt("模头厂内初调")
								sel_opt("模头厂外精调")
								sel_opt("模头预验收或寄样")
								sel_opt("模头来厂验收")
							End If
							if not isnull(rs("dxtsr")) and isnull(rs("dxtsjs")) then
								sel_opt("结束定型调试")
								sel_opt("定型厂内初调")
								sel_opt("定型厂外精调")
								sel_opt("定型预验收或寄样")
								sel_opt("定型来厂验收")
							End If
							if isnull(rs("mttsxxzlr")) and not isnull(rs("mttsjs")) then sel_opt("开始模头调试信息整理")
							if isnull(rs("dxtsxxzlr")) and not isnull(rs("dxtsjs")) then sel_opt("开始定型调试信息整理")
							if isnull(rs("mttsxxzlr")) and not isnull(rs("mttsjs")) and isnull(rs("dxtsxxzlr")) and not isnull(rs("dxtsjs")) then sel_opt("开始全套调试信息整理")

							if not isnull(rs("mttsxxzlr")) and isnull(rs("mttsxxzljs")) then sel_opt("结束模头调试信息整理")
							if not isnull(rs("dxtsxxzlr")) and isnull(rs("dxtsxxzljs")) then sel_opt("结束定型调试信息整理")
							if (rs("mttsxxzlr")=rs("dxtsxxzlr")) and isnull(rs("mttsxxzljs")) and isnull(rs("dxtsxxzljs")) then sel_opt("结束全套调试信息整理")
						case "模头"
							if isnull(rs("mttsr")) and not isnull(rs("mttsdjs")) then sel_opt("开始模头调试")
							if not isnull(rs("mttsr")) and isnull(rs("mttsjs")) then
								sel_opt("结束模头调试")
								sel_opt("模头厂内初调")
								sel_opt("模头厂外精调")
								sel_opt("模头预验收或寄样")
								sel_opt("模头来厂验收")
							End If
							if isnull(rs("mttsxxzlr")) and not isnull(rs("mttsjs")) then sel_opt("开始模头调试信息整理")
							if not isnull(rs("mttsxxzlr")) and isnull(rs("mttsxxzljs")) then sel_opt("结束模头调试信息整理")

						case "定型"
							if isnull(rs("dxtsr")) and not isnull(rs("dxtsdjs")) then sel_opt("开始定型调试")
							if not isnull(rs("dxtsr")) and isnull(rs("dxtsjs")) then
								sel_opt("结束定型调试")
								sel_opt("定型厂内初调")
								sel_opt("定型厂外精调")
								sel_opt("定型预验收或寄样")
								sel_opt("定型来厂验收")
							End If
							if isnull(rs("dxtsxxzlr")) and not isnull(rs("dxtsjs")) then sel_opt("开始定型调试信息整理")
							if not isnull(rs("dxtsxxzlr")) and isnull(rs("dxtsxxzljs")) then sel_opt("结束定型调试信息整理")
						case else
							response.write(rs("mjxx"))
					end select
				%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type=submit value=" 分配 "></td>
    </tr>
    <input type=hidden name=lsh value=<%=rs("lsh")%>>
  </form>
</table>
<%
		call atask_js()
	End If
End Function

Function atask_nofinished()			'具有分配权限的未完成的调试任务
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter,TotalCount, sqlorder
	absPageNum = 0
	RecordPerPage = 40
	iCounter = 1
	sqlorder = " order by jhjssj"
	If LCase(strOrder) = "ddh" Then sqlorder = " order by ddh desc, lsh desc"
	If LCase(strOrder) = "lsh" Then sqlorder = " order by lsh desc"

	strSql="select * from [mtask] where not(isnull(fsjs)) and not(mjjs) and ((mjxx='全套' and (isnull(mttsdjs) or isnull(dxtsdjs)) and ([group]="&session("userGroup")&" Or zz='"&Session("userName")&"' Or jgzz='"&Session("userName")&"' Or sjzz='"&Session("userName")&"')) or (mjxx='模头' and isnull(mttsdjs) and ([group]="&session("userGroup")&" Or zz='"&Session("userName")&"' Or jgzz='"&Session("userName")&"' Or sjzz='"&Session("userName")&"')) or (mjxx='定型' and isnull(dxtsdjs) and ([group]="&session("userGroup")&" Or zz='"&Session("userName")&"' Or jgzz='"&Session("userName")&"' Or sjzz='"&Session("userName")&"')) or ((mjxx='全套' and (not isnull(mttsdjs)) and (not isnull(dxtsdjs)) and "&session("userGroup")&"=5) or (mjxx='模头' and (not isnull(mttsdjs)) and "&session("userGroup")&"=5) or (mjxx='定型' and (not isnull(dxtsdjs)) and "&session("userGroup")&"=5) ))" & sqlorder
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,1,3
	If (Rs.Eof Or Rs.Bof) Then
		Call TbTopic("暂时没有待分配辅助任务!") : Exit Function
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
%>
<%Call TbTopic("您共有 "& rs.recordcount &" 套待分配的辅助任务")%>
<table width="95%" cellspacing=0 cellpadding=2 class=xtable align="center">
  <tr>
    <th class="th">id</th>
    <th class="th">订单号</th>
    <th class="th">流水号</th>
    <th class="th">单位名称</th>
    <th class="th">断面名称</th>
    <th class="th">组长</th>
    <th class="th">任务内容</th>
    <th class=th width=*>调试单</th>
    <th class=th width=*>调试</th>
    <th class=th width=*>调试整理</th>
  </tr>
  <% for absrecordnum = 1 to recordperpage %>
  <tr>
    <td class="ctd"><%=icounter %></td>
    <td class="ctd"><%= rs("ddh")%></td>
    <td class="ctd"><a href=atask_assign.asp?s_lsh=<%=rs("lsh")%>><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("dwmc")%></td>
    <td class="ctd"><%=rs("dmmc")%></td>
    <td class="ctd"><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(j)、"&rs("sjzz")&"(s)")%></td>
    <td class="ctd"><%= rs("mjxx") & rs("rwlr") %></td>
    <%select case rs("mjxx")%>
    <%case "全套"%>
    <td class="ctd"><%call distd(rs("mttsdks"),rs("mttsdjs"),-20,rs)%>
      <%=DATEADD("d",20,rs("jhjssj"))%>
      <%call distd(rs("dxtsdks"),rs("dxtsdjs"),-20,rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsks"),rs("mttsjs"),rs)%>
      <%call distd2(rs("dxtsks"),rs("dxtsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
      <%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <%case "模头"%>
    <td class="ctd"><%call distd(rs("mttsdks"),rs("mttsdjs"),-20,rs)%>
      <%=DATEADD("d",20,rs("jhjssj"))%> </td>
    <td class="ctd"><%call distd2(rs("mttsks"),rs("mttsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
    </td>
    <%case "定型"%>
    <td class="ctd"><%=DATEADD("d",20,rs("jhjssj"))%>
      <%call distd(rs("dxtsdks"),rs("dxtsdjs"),-20,rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("dxtsks"),rs("dxtsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <%end select%>
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
      <select name="ipage" onchange='location.href("<%=Request.ServerVariables("script_name")%>?ipage=" + this.value+"");'>
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
	rs.close
end function

function sel_opt(str)
%>
<option value="<%=str%>"><%=str%></option>
<%
end function

function atask_js()
%>
<script language="javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='开始')
				objdoc.form_assign.zrr.disabled=false;
			else
				objdoc.form_assign.zrr.disabled=true;
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
%>
