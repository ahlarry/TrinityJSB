<%
'15:08 2007-1-6-星期六
Rem 显示任务信息的代码(因为要在多处使用所以放在此文件内)
Function mtask_fewinfo(rs)
%>
<%Call TbTopic("挤出模具厂挤出模设计任务书")%>

<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4"><b>■合同信息■</b></td>
  </tr>
  <tr>
    <td class="rtd" width="15%">订单号</td>
    <td class="ltd" width="35%"><%=rs("ddh")%></td>
    <td class="rtd" width="15%">流水号</td>
    <td width="*" class="ltd"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
  </tr>
  <tr>
    <td class="rtd">客户名称</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">断面名称</td>
    <td class="ltd"><%=rs("dmmc")%></td>
  </tr>
  <tr>
    <td class="rtd">模号</td>
    <td class="ltd"><%=rs("mh")%></td>
    <td class="rtd">设备厂家</td>
    <td class="ltd"><%=rs("sbcj")%></td>
  </tr>
  <tr>
    <td class="rtd">任务内容</td>
    <td class="ltd">
    <%
    If IsNull(rs("mtrw")) and IsNull(rs("dxrw")) Then
    	Response.Write(rs("mjxx") & rs("rwlr"))
    else
    	If Rs("mtrw")<>"" Then Response.Write("模头"&rs("mtrw")) End If
    	If Rs("dxrw")<>"" Then Response.Write("定型"&rs("dxrw")) End If
    End If
    %>
   </td>
    <td class="rtd">计划结束时间</td>
    <td class="ltd"><%=xjDate(rs("jhjssj"),1)%></td>
  </tr>
</table>
<%
End Function
'主要信息
Function mtask_muchinfo(rs)
%>
<%Call TbTopic("挤出模具厂挤出模设计任务书")%>
<%If Chkable("1,3") Then%>
<a href="mtask_print.asp?s_lsh=<%=rs("lsh")%>">打印任务书</a>
<%End If%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="6"><b>合同信息</b></td>
  </tr>
  <tr>
    <td class="rtd" width="13%">订单号</td>
    <td class="ltd"><%=rs("ddh")%></td>
    <td class="rtd" width="13%">流水号</td>
    <td class="ltd" width="*"><a href="SreacDwg.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="rtd" width="13%">模号</td>
    <td class="ltd" width="*"><%=rs("mh")%></td>
  </tr>
  <tr>
    <td class="rtd">客户名称</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">断面名称</td>
    <td class="ltd"><%=rs("dmmc")%></td>
    <td class="rtd">模具材料</td>
    <td class="ltd"><%=rs("mjcl")%></td>
  </tr>
  <tr>
    <td class="rtd">设备厂家</td>
    <td class="ltd"><%=rs("sbcj")%></td>
    <td class="rtd">水接头数量</td>
    <td class="ltd"><%=rs("sjtsl")%></td>
    <td class="rtd">气接头数量</td>
    <td class="ltd"><%=rs("qjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">挤出机型号</td>
    <td class="ltd"><%=rs("jcjxh")%></td>
    <td class="rtd">挤出方向</td>
    <td class="ltd"><%=rs("jcfx")%>&nbsp;</td>
    <td class="rtd">牵引速度</td>
    <td class="ltd"><%=rs("qysd")%> 米/分(m/min)</td>
  </tr>
  <tr>
    <td class="rtd">配加热板</td>
    <td class="ltd"><%if rs("pjrb") then%>
      是
      <%else%>
      否
      <%end if%></td>
    <td class="rtd">加热板信息</td>
    <td class="ltd">相数:<%=rs("jrbxs")%> 材质:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
    <td class="rtd">腔数</td>
    <td class="ltd"><%=rs("qs")%>腔</td>
  </tr>
</table>
<%
End Function

'工艺信息
Function mtask_technicsinfo(rs)
Dim ssgjf, qbfgjf, qgjf, hgjf,gjxx
ssgjf=NullToNum(Rs("ssgj"))
qbfgjf=NullToNum(Rs("qbfgj"))
qgjf=NullToNum(Rs("qgj"))
hgjf=NullToNum(Rs("hgj"))
gjxx=""
select case ssgjf&qbfgjf&qgjf&hgjf
Case "0000"			'兼容08版共挤计分模式
	If Rs("gjzf")>0 and Rs("gjfs")=1 Then
 		gjxx="双色共挤"
 	Elseif Rs("gjzf")>0 and Rs("gjfs")=2 Then
  		gjxx="全包覆共挤"
  	Elseif Rs("gjfs")=3 and Rs("qhgj")=1 Then
  		gjxx="软硬前共挤"
  	Elseif Rs("gjfs")=3 and Rs("qhgj")=2 Then
  		gjxx="软硬后共挤"
  	Else
  		gjxx="/"
  	End If
Case Else		'09版共挤计分模式
	If ssgjf<>0 Then gjxx="双色共挤"
	If qbfgjf<>0 Then gjxx=gjxx &" 全包覆共挤"
	If qgjf<>0 Then gjxx=gjxx &" 软硬前共挤"
	If hgjf<>0 Then gjxx=gjxx &" 软硬后共挤"
end select
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="6"><b>模具信息</b></td>
  </tr>
  <tr>
    <td class="rtd"  width="13%">任务内容</td>
    <td class="ltd">
    <%
    If IsNull(rs("mtrw")) and IsNull(rs("dxrw")) Then
    	Response.Write(rs("mjxx") & rs("rwlr"))
    else
    	If Rs("mtrw")<>"" Then Response.Write("模头"&rs("mtrw")) End If
    	If Rs("dxrw")<>"" Then Response.Write("定型"&rs("dxrw")) End If
    End If
    %>
   </td>
        <td class="rtd"  width="13%">厂内调试</td>
    <td class="ltd"  width="*"><%if rs("cnts") then%>
      是
      <%else%>
      &nbsp;/
      <%end if%></td>
    <td class="rtd"  width="13%">调试类别</td>
    <% If Rs("cnts") Then%>
    <%If Not(isnull(Rs("tslb"))) Then%>
    <td class="ltd"  width="*"><a href="mtest_display.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("tslb")%></a></td>
    <%Else%>
    <td class="ltd"  width="*">&nbsp;/</td>
    <%End If%>
    <%Else%>
    <%If Rs("beit") Then%>
    <td class="ltd"  width="*">北调</td>
    <%Else%>
    <td class="ltd"  width="*">&nbsp;/</td>
    <%End If%>
    <%End If%>
  </tr>
  <tr>
    <td class="rtd">模头结构</td>
    <td class="ltd"><%if IsNull(rs("mtjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("mtjg"))
    End if%></td>
    <td class="rtd">定型结构</td>
    <td class="ltd"><%if IsNull(rs("dxjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("dxjg"))
    End if%></td>
    <td class="rtd">水箱结构</td>
    <td class="ltd"><%if IsNull(rs("sxjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("sxjg"))
    End if%></td>
  </tr>
  <tr>
    <td class="rtd">定型切割</td>
    <td class="ltd"><%if IsNull(rs("dxqg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("dxqg"))
    End if%></td>
    <td class="rtd">模头连接尺寸</td>
    <td class="ltd"><%=rs("mtljcc")%></td>
    <td class="rtd">热电偶规格</td>
    <td class="ltd"><%=rs("rdogg")%></td>
  </tr>
  <tr>
    <td class="rtd">共挤类型</td>
    <td class="ltd"><%=Trim(gjxx)%></td>
    <td class="rtd">共挤连接尺寸</td>
    <td class="ltd"><%=rs("gjljcc")%>&nbsp;</td>
    <td class="rtd">型材壁厚</td>
    <td class="ltd"><%=Rs("xcbh")%>毫米</td>
  </tr>
  <tr>
  </tr>
  </table>
  <%Response.Write(XjLine(5,web_info(8),""))%>
  <table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8" alt="<%=DisFzInfo(Rs)%>"><b>其他信息</b></td>
  </tr>
  <tr>
    <td class="rtd" >评审记录</td>
    <td class="ltd" colspan="7" height="120" valign="top"><%=xjweb.HtmlToCode(rs("psjl"))%></td>
  </tr>
  <tr>
    <td class="rtd">备注</td>
    <td class="ltd" colspan="7" height="150" valign="top"><%=xjweb.HtmlToCode(rs("bz"))%></td>
  </tr>
  <tr>
    <td class="rtd">计划开始</td>
    <%If rs("jhkssj")<>"" Then%>
    <td class="ltd"><%=XjDate(rs("jhkssj"),3)%></td>
    <%else%>
    <td class="ltd" >&nbsp;/</td>
    <%End If%>
    <td class="rtd">计划结构结束</td>
    <td class="ltd" width="12%"><%=XjDate(rs("jhjgsj"),3)%></td>
    <td class="rtd">计划全套结束</td>
    <td class="ltd"><%=XjDate(rs("jhjssj"),3)%></td>
    <td class="rtd">实际结束</td>
    <td class="ltd" width="12%"><%=XjDate(rs("sjjssj"),3)%></td>
  </tr>
  <tr>
    <td class="rtd">组长</td>
    <td colspan="3" class="ltd"><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(结构)、"&rs("sjzz")&"(设计)")%></td>
    <td class="rtd">技术代表</td>
    <td colspan="3" class="ltd"><%=rs("jsdb")%></td>
  </tr>
</table>
<% End Function
'全部信息
Function mtask_allinfo(rs)
%>
<%Call TbTopic("挤出模具厂挤出模设计任务书")%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4"><b>合同信息</b></td>
  </tr>
  <tr>
    <td class="rtd" width="13%">订单号</td>
    <td class="ltd" width="37%"><%=rs("ddh")%></td>
    <td class="rtd" width="13%">流水号</td>
    <td class="ltd" width="*"><%=rs("lsh")%></td>
  </tr>
  <tr>
    <td class="rtd">客户名称</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">断面名称</td>
    <td class="ltd"><%=rs("dmmc")%></td>
  </tr>
  <tr>
    <td class="rtd">模号</td>
    <td class="ltd"><%=rs("mh")%></td>
    <td class="rtd">模具材料</td>
    <td class="ltd"><%=rs("mjcl")%></td>
  </tr>
  <tr>
    <td class="rtd">设备厂家</td>
    <td class="ltd"><%=rs("sbcj")%></td>
    <td class="rtd">水接头数量</td>
    <td class="ltd"><%=rs("sjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">挤出机型号</td>
    <td class="ltd"><%=rs("jcjxh")%></td>
    <td class="rtd">气接头数量</td>
    <td class="ltd"><%=rs("qjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">配加热板</td>
    <td class="ltd"><%if rs("pjrb") then%>
      是
      <%else%>
      否
      <%end if%></td>
    <td class="rtd">腔数</td>
    <td class="ltd"><%=rs("qs")%>腔</td>
  </tr>
  <tr>
    <td class="rtd">加热板信息</td>
    <td class="ltd">相数:<%=rs("jrbxs")%> 材质:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
    <td class="rtd">牵引速度</td>
    <td class="ltd"><%=rs("qysd")%> 米/分(m/min)</td>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4"><b>模具信息</b></td>
  </tr>
  <tr>
    <td class="rtd">定型切割</td>
    <td class="ltd"><%=rs("dxqg")%>&nbsp;</td>
    <td class="rtd">任务内容</td>
    <td class="ltd">
    <%
    If IsNull(rs("mtrw")) and IsNull(rs("dxrw")) Then
    	Response.Write(rs("mjxx") & rs("rwlr"))
    else
    	If Rs("mtrw")<>"" Then Response.Write("模头"&rs("mtrw")) End If
    	If Rs("dxrw")<>"" Then Response.Write("定型"&rs("dxrw")) End If
    End If
    %>
   </td>
  </tr>
  <tr>
    <td class="rtd">定型结构</td>
    <td class="ltd"><%=rs("dxjg")%>&nbsp;</td>
    <td class="rtd">模头连接尺寸</td>
    <td class="ltd"><%=rs("mtljcc")%></td>
  </tr>
  <tr>
    <td class="rtd">水箱结构</td>
    <td class="ltd"><%=rs("sxjg")%>&nbsp;</td>
    <td class="rtd">热电偶规格</td>
    <td class="ltd"><%=rs("rdogg")%></td>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4" alt="<%=DisFzInfo(Rs)%>"><b>其他信息</b></td>
  </tr>
  <tr>
    <td class="rtd" >评审记录</td>
    <td class="ltd" colspan="3" height="120" valign="top"><%=xjweb.HtmlToCode(rs("psjl"))%></td>
  </tr>
  <tr>
    <td class="rtd" >备注</td>
    <td class="ltd" colspan="3" height="150" valign="top"><%=xjweb.HtmlToCode(rs("bz"))%></td>
  </tr>
  <tr>
    <td class="rtd">计划结束时间</td>
    <td class="ltd"><%=XjDate(rs("jhjssj"),1)%></td>
    <td class="rtd">实际结束时间;</td>
    <td class="ltd"><%=XjDate(rs("sjjssj"),1)%></td>
  </tr>
  <tr>
    <td class="rtd">组长</td>
    <td colspan="3" class="ltd"><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(结构)、"&rs("sjzz")&"(设计)")%></td>
    <td class="rtd">技术代表</td>
    <td class="ltd"><%=rs("jsdb")%></td>
  </tr>
</table>
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call mtask_userinfo(rs)%>
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call atask_userinfo(rs)%>
<%
End Function

Function mtask_userinfo(rs)
Dim strgy
strgy=""
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <%
			select case rs("mjxx") & rs("rwlr")
				case "全套设计"
				 If ((rs("gjfs")=3) and (rs("qhgj")=2)) or NullToNum(Rs("hgj"))<>0 Then%>
  <tr>
    <td class="ctd" width="10%">后共挤结构</td>
    <td class="ctd" width="*"><%=rs("gjjgr")%>&nbsp;</td>
    <td class="ctd" width="13%">后共挤结构确认</td>
    <td class="ctd" width="*"><%=rs("gjjgshr")%>&nbsp;</td>
    <td class="ctd" width="10%">后共挤设计</td>
    <td class="ctd" width="*"><%=rs("gjsjr")%>&nbsp;</td>
    <%If rs("gjshr")<>"" Then%>
    <td class="ctd" width="10%">后共挤审核</td>
    <td class="ctd" width="*"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="*" colspan="4">&nbsp;</td>
    <%else%>
    <td class="ctd" width="13%">后共挤设计审核</td>
    <td class="ctd" width="*" colspan="3"><%=rs("gjsjshr")%>&nbsp;</td>
    <%End If%>
  </tr>
  <% End If%>
  <%if (not isnull(rs("mtjgr"))) and (not isnull(rs("dxjgr"))) and rs("mtjgr")=rs("dxjgr") then%>
  <tr>
    <td class="ctd" width="10%" rowspan="2">模具结构</td>
    <%else%>
    <td class="ctd" width="10%">模头结构</td>
    <%end if%>
    <%if rs("mtjgr")=rs("dxjgr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtjgr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtjgr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtjgshr")=rs("dxjgshr") then%>
    <td class="ctd" width="11%" rowspan="2">结构确认</td>
    <%else%>
    <td class="ctd" width="11%">模头结构确认</td>
    <%end if%>
    <%if rs("mtjgshr")=rs("dxjgshr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtjgshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtjgshr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="10%" rowspan="2">模具设计</td>
    <%else%>
    <td class="ctd" width="10%">模头设计</td>
    <%end if%>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtsjr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtsjr")%>&nbsp;</td>
    <%end if%>
    <%If (not isnull(rs("mtshr"))) or (not isnull(rs("dxshr"))) Then%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="10%" rowspan="2">模具审核</td>
    <%else%>
    <td class="ctd" width="10%">模头审核</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtshr")%>&nbsp;</td>
    <%end if%>
    <%else%>
    <%if rs("mtsjshr")=rs("dxsjshr") then%>
    <td class="ctd" width="11%" rowspan="2">设计审核</td>
    <%else%>
    <td class="ctd" width="11%">模头设计审核</td>
    <%end if%>
    <%if rs("mtsjshr")=rs("dxsjshr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtsjshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtsjshr")%>&nbsp;</td>
    <%end if%>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="8%" rowspan="2">模具BOM</td>
    <%else%>
    <td class="ctd" width="10%">模头BOM</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtbomr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <tr>
    <%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
    <td class="ctd" width="10%">定型结构</td>
    <%end if%>
    <%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
    <td class="ctd" width="8%"><%=rs("dxjgr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtjgshr")) or isnull(rs("dxjgshr")) or (rs("mtjgshr")<>rs("dxjgshr")) then%>
    <td class="ctd" width="11%">定型结构确认</td>
    <%end if%>
    <%if isnull(rs("mtjgshr")) or isnull(rs("dxjgshr")) or (rs("mtjgshr")<>rs("dxjgshr")) then%>
    <td class="ctd" width="8%"><%=rs("dxjgshr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="10%">定型设计</td>
    <%end if%>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="8%"><%=rs("dxsjr")%>&nbsp;</td>
    <%end if%>
    <%If (not isnull(rs("mtshr"))) or (not isnull(rs("dxshr"))) Then%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="10%">定型审核</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="8%"><%=rs("dxshr")%>&nbsp;</td>
    <%end if%>
    <%else%>
    <%if isnull(rs("mtsjshr")) or isnull(rs("dxsjshr")) or (rs("mtsjshr")<>rs("dxsjshr")) then%>
    <td class="ctd" width="11%">定型设计审核</td>
    <%end if%>
    <%if isnull(rs("mtsjshr")) or isnull(rs("dxsjshr")) or (rs("mtsjshr")<>rs("dxsjshr")) then%>
    <td class="ctd" width="8%"><%=rs("dxsjshr")%>&nbsp;</td>
    <%end if%>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="10%">定型BOM</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="8%"><%=rs("dxbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <%case "全套复改"	%>
  <%If Rs("gjsjr")<>"" Then%>
  <tr>
    <td class="ctd" width="10%" colspan="2">　</td>
    <td class="ctd" width="10%">共挤复改</td>
    <td class="ctd" width="15%" colspan="2"><%=rs("gjsjr")%>&nbsp;</td>
    <td class="ctd" width="10%">共挤审核</td>
    <td class="ctd" width="15%" colspan="2"><%=rs("gjshr")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="10%" rowspan="2">　</td>
    <td class="ctd" width="15%" rowspan="2">　</td>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="10%" rowspan="2">模具复改</td>
    <%else%>
    <td class="ctd" width="10%">模头复改</td>
    <%end if%>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="15%" rowspan="2"><%=rs("mtsjr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="15%"><%=rs("mtsjr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="10%" rowspan="2">模具审核</td>
    <%else%>
    <td class="ctd" width="10%">模头审核</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="15%" rowspan="2"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="10%" rowspan="2">模具BOM</td>
    <%else%>
    <td class="ctd" width="10%">模头BOM</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="*" rowspan="2"><%=rs("mtbomr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <tr>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="10%">定型复改</td>
    <%end if%>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="15%"><%=rs("dxsjr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="10%">定型审核</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="10%">定型BOM</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <%	case "全套复查"	%>
  <%If ((rs("gjfs")=3) and (rs("qhgj")=2)) or NullToNum(Rs("hgj"))<>0  Then%>
  <tr>
    <td class="ctd" width="50%">&nbsp;</td>
    <td class="ctd" width="10%">共挤复查</td>
    <td class="ctd" width="15%" colspan="2"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width=* colspan="2">&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="50%">&nbsp;</td>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="10%" rowspan="2">模具复查</td>
    <%else%>
    <td class="ctd" width="10%">模头复查</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="15%" rowspan="2"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="10%" rowspan="2">模具BOM</td>
    <%else%>
    <td class="ctd" width="10%">模头BOM</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="*" rowspan="2"><%=rs("mtbomr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <tr>
    <td class="ctd" width="50%">&nbsp;</td>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="10%">定型复查</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="10%">定型BOM</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%end if%>
    <%	case "模头设计"	%>
    <td class="ctd" width="10%">模头结构</td>
    <td class="ctd" width="10%"><%=rs("mtjgr")%>&nbsp;</td>
    <td class="ctd" width="11%">模头结构确认</td>
    <td class="ctd" width="10%"><%=rs("mtjgshr")%>&nbsp;</td>
    <td class="ctd" width="10%">模头设计</td>
    <td class="ctd" width="10%"><%=rs("mtsjr")%>&nbsp;</td>
    <%If (not(isnull(Rs("mtshr")))) Then%>
    <td class="ctd" width="10%">模头审核</td>
    <td class="ctd" width="10%"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="11%">模头设计审核</td>
    <td class="ctd" width="10%"><%=rs("mtsjshr")%>&nbsp;</td>
    <%End If%>
    <td class="ctd" width="10%">模头BOM</td>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%	case "模头复改"	%>
    <td class="ctd" width="10%">　</td>
    <td class="ctd" width="15%">　</td>
    <td class="ctd" width="10%">模头复改</td>
    <td class="ctd" width="15%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="10%">模头审核</td>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="10%">模头BOM</td>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%	case "模头复查"		%>
    <td class="ctd" width="10%">　</td>
    <td class="ctd" width="15%">　</td>
    <td class="ctd" width="10%">　</td>
    <td class="ctd" width="15%">　</td>
    <td class="ctd" width="10%">模头复查</td>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="10%">模头BOM</td>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%	case "定型设计"		%>
    <td class="ctd" width="10%">定型结构</td>
    <td class="ctd" width="10%"><%=rs("dxjgr")%>&nbsp;</td>
    <td class="ctd" width="11%">定型结构确认</td>
    <td class="ctd" width="10%"><%=rs("dxjgshr")%>&nbsp;</td>
    <td class="ctd" width="10%">定型设计</td>
    <td class="ctd" width="10%"><%=rs("dxsjr")%>&nbsp;</td>
    <%If (not(isnull(Rs("dxshr")))) Then%>
    <td class="ctd" width="10%">定型审核</td>
    <td class="ctd" width="10%"><%=rs("dxshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="11%">定型设计审核</td>
    <td class="ctd" width="10%"><%=rs("dxsjshr")%>&nbsp;</td>
    <%End If%>
    <td class="ctd" width="10%">定型BOM</td>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%	case "定型复改"		%>
    <td class="ctd" width="10%">　</td>
    <td class="ctd" width="15%">　</td>
    <td class="ctd" width="10%">定型复改</td>
    <td class="ctd" width="15%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="10%">定型审核</td>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="10%">定型BOM</td>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%		case "定型复查"		%>
    <td class="ctd" width="10%">　</td>
    <td class="ctd" width="15%">　</td>
    <td class="ctd" width="10%">　</td>
    <td class="ctd" width="15%">　</td>
    <td class="ctd" width="10%">定型复查</td>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="10%">定型BOM</td>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%
				case else
				response.write(rs("mjxx") & rs("rwlr"))
			end select
		%>
  </tr>
</table>
<%
Response.Write(XjLine(5, "100%", ""))
If not(isNull(rs("gysjr"))) Then
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">工艺设计</td>
    <td class="ctd" width="35%" ><%=rs("gysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">工艺审核</td>
    <td class="ctd" width="35%" ><%=rs("gyshr")%>&nbsp;</td>
  </tr>
</table>
<%
else
select case rs("mjxx")
	case "模头"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">模头工艺设计</td>
    <td class="ctd" width="35%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">模头工艺审核</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
  </tr>
</table>
<%
	case "定型"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">定型工艺设计</td>
    <td class="ctd" width="35%"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">定型工艺审核</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
  </tr>
</table>
<%
	case else
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">模头工艺设计</td>
    <td class="ctd" width="35%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">模头工艺审核</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型工艺设计</td>
    <td class="ctd" width="35%"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">定型工艺审核</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
  </tr>
  <%If Rs("gjgysjr")<>"" Then%>
  <tr>
    <td class="ctd" width="10%">共挤工艺设计</td>
    <td class="ctd" width="35%"><%=rs("gjgysjr")%>&nbsp;</td>
    <td class="ctd" width="10%">共挤工艺审核</td>
    <td class="ctd"><%=rs("gjgyshr")%>&nbsp;</td>
  </tr>
  <%End If%>
</table>
<%
End select
End If
end function

function atask_userinfo(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <%
			select case rs("mjxx")
				case "全套"
				%>
    <%if (not isnull(rs("mttsdr"))) and (not isnull(rs("dxtsdr"))) and (rs("mttsdr")=rs("dxtsdr")) then%>
    <td class="ctd" width="15%" rowspan="2">模具调试单</td>
    <%else%>
    <td class="ctd" width="15%">模头调试单</td>
    <%end if%>
    <%if rs("mttsdr")=rs("dxtsdr") then%>
    <td class="ctd" width="10%" rowspan="2"><%=rs("mttsdr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="10%"><%=rs("mttsdr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mttsr")=rs("dxtsr") then%>
    <td class="ctd" width="15%" rowspan="2">模具调试</td>
    <%else%>
    <td class="ctd" width="15%">模头调试</td>
    <%end if%>
    <%if rs("mttsr")=rs("dxtsr") then%>
    <td class="ctd" width="10%" rowspan="2"><%=rs("mttsr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="10%"><%=rs("mttsr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
    <td class="ctd" width="15%" rowspan="2">模具调试信息整理</td>
    <%else%>
    <td class="ctd" width="15%">模头调试信息整理</td>
    <%end if%>
    <%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
    <td class="ctd" width="10%" rowspan="2"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="10%"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <%end if%>
    <td class="ctd" width="15%" rowspan="2">齐套信息整理</td>
    <td class="ctd" width="10%" rowspan="2"><%=rs("xtxxzlr")%>&nbsp;</td>
  </tr>
  <tr>
    <%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
    <td class="ctd" width="15%">定型调试单</td>
    <%end if%>
    <%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
    <td class="ctd" width="10%"><%=rs("dxtsdr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
    <td class="ctd" width="15%">定型调试</td>
    <%end if%>
    <%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
    <td class="ctd" width="10%"><%=rs("dxtsr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
    <td class="ctd" width="15%">定型调试信息整理</td>
    <%end if%>
    <%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
    <td class="ctd" width="10%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <%end if%>
    <%
				case "模头"
				%>
    <td class="ctd" width="15%">模头调试单</td>
    <td class="ctd" width="15%"><%=rs("mttsdr")%>&nbsp;</td>
    <td class="ctd" width="15%">模头调试</td>
    <td class="ctd" width="15%"><%=rs("mttsr")%>&nbsp;</td>
    <td class="ctd" width="15%">模头调试信息整理</td>
    <td class="ctd" width="*"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="15%">齐套信息整理</td>
    <td class="ctd" width="*"><%=rs("xtxxzlr")%>&nbsp;</td>
    <%
				case "定型"
				%>
    <td class="ctd" width="15%">定型调试单</td>
    <td class="ctd" width="15%"><%=rs("dxtsdr")%>&nbsp;</td>
    <td class="ctd" width="15%">定型调试</td>
    <td class="ctd" width="15%"><%=rs("dxtsr")%>&nbsp;</td>
    <td class="ctd" width="15%">定型调试信息整理</td>
    <td class="ctd" width="*"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="15%">齐套信息整理</td>
    <td class="ctd" width="*"><%=rs("xtxxzlr")%>&nbsp;</td>
    <%
				case else
				response.write(rs("mjxx"))
			end select
		%>
  </tr>
</table>
<%
end function


function mtask_alluserinfo(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <%select case rs("mjxx") & rs("rwlr")%>
  <%case "全套设计"%>
  <tr>
    <td class="ctd" width="15%">模头结构</td>
    <td class="ctd" width="9%"><%=rs("mtjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构开始</td>
    <td class="ctd" width="20%"><%=rs("mtjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构结束</td>
    <td class="ctd" width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">模头结构确认</td>
    <td class="ctd" width="9%"><%=rs("mtjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构确认开始</td>
    <td class="ctd" width="20%"><%=rs("mtjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构确认结束</td>
    <td class="ctd" width="20%"><%=rs("mtjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">模头设计</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计开始</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计结束</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">模头审核</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核开始</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核结束</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">模头设计审核</td>
    <td class="ctd" width="9%"><%=rs("mtsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计审核开始</td>
    <td class="ctd" width="20%"><%=rs("mtsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计审核结束</td>
    <td class="ctd" width="20%"><%=rs("mtsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">模头BOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM开始</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM结束</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型结构</td>
    <td class="ctd" width="9%"><%=rs("dxjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构开始</td>
    <td class="ctd" width="20%"><%=rs("dxjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构结束</td>
    <td class="ctd" width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">定型结构确认</td>
    <td class="ctd" width="9%"><%=rs("dxjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构确认开始</td>
    <td class="ctd" width="20%"><%=rs("dxjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构确认结束</td>
    <td class="ctd" width="20%"><%=rs("dxjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">定型设计</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计开始</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计结束</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">定型审核</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核开始</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核结束</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">定型设计审核</td>
    <td class="ctd" width="9%"><%=rs("dxsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计审核开始</td>
    <td class="ctd" width="20%"><%=rs("dxsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计审核结束</td>
    <td class="ctd" width="20%"><%=rs("dxsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">定型BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM开始</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM结束</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <% If (not isnull(rs("gjjgr"))) Then%>
  <tr>
    <td class="ctd" width="15%">后共挤结构</td>
    <td class="ctd" width="9%"><%=rs("gjjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">后共挤结构开始</td>
    <td class="ctd" width="20%"><%=rs("gjjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">后共挤结构结束</td>
    <td class="ctd" width="20%"><%=rs("gjjgjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">后共挤设计</td>
    <td class="ctd" width="9%"><%=rs("gjsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">后共挤设计开始</td>
    <td class="ctd" width="20%"><%=rs("gjsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">后共挤设计结束</td>
    <td class="ctd" width="20%"><%=rs("gjsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">后共挤审核</td>
    <td class="ctd" width="9%"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">后共挤审核开始</td>
    <td class="ctd" width="20%"><%=rs("gjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">后共挤审核结束</td>
    <td class="ctd" width="20%"><%=rs("gjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <%case "模头设计"%>
  <tr>
    <td class="ctd" width="15%">模头结构</td>
    <td class="ctd" width="9%"><%=rs("mtjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构开始</td>
    <td class="ctd" width="20%"><%=rs("mtjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构结束</td>
    <td class="ctd" width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">模头结构确认</td>
    <td class="ctd" width="9%"><%=rs("mtjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构确认开始</td>
    <td class="ctd" width="20%"><%=rs("mtjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头结构确认结束</td>
    <td class="ctd" width="20%"><%=rs("mtjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">模头设计</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计开始</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计结束</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">模头审核</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核开始</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核结束</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">模头设计审核</td>
    <td class="ctd" width="9%"><%=rs("mtsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计审核开始</td>
    <td class="ctd" width="20%"><%=rs("mtsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头设计审核结束</td>
    <td class="ctd" width="20%"><%=rs("mtsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">模头BOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM开始</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM结束</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <%case "定型设计"%>
  <tr>
    <td class="ctd" width="15%">定型结构</td>
    <td class="ctd" width="9%"><%=rs("dxjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构开始</td>
    <td class="ctd" width="20%"><%=rs("dxjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构结束</td>
    <td class="ctd" width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">定型结构确认</td>
    <td class="ctd" width="9%"><%=rs("dxjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构确认开始</td>
    <td class="ctd" width="20%"><%=rs("dxjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型结构确认结束</td>
    <td class="ctd" width="20%"><%=rs("dxjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">定型设计</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计开始</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计结束</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">定型审核</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核开始</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核结束</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">定型设计审核</td>
    <td class="ctd" width="9%"><%=rs("dxsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计审核开始</td>
    <td class="ctd" width="20%"><%=rs("dxsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型设计审核结束</td>
    <td class="ctd" width="20%"><%=rs("dxsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">定型BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM开始</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM结束</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <%case "全套复改"%>
  <tr>
    <td class="ctd" width="15%">模头复改</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复改开始</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复改结束</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头审核</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核开始</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核结束</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头BOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM开始</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM结束</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型复改</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复改开始</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复改结束</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型审核</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核开始</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核结束</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM开始</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM结束</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <% If (not isnull(rs("gjsjr"))) Then%>
  <tr>
    <td class="ctd" width="15%">共挤复改</td>
    <td class="ctd" width="9%"><%=rs("gjsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">共挤复改开始</td>
    <td class="ctd" width="20%"><%=rs("gjsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">共挤复改结束</td>
    <td class="ctd" width="20%"><%=rs("gjsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">共挤审核</td>
    <td class="ctd" width="9%"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">共挤审核开始</td>
    <td class="ctd" width="20%"><%=rs("gjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">共挤审核结束</td>
    <td class="ctd" width="20%"><%=rs("gjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <%case "模头复改"%>
  <tr>
    <td class="ctd" width="15%">模头复改</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复改开始</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复改结束</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头复改</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复改开始</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复改结束</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头审核</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核开始</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头审核结束</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头BOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM开始</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM结束</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <%case "定型复改"%>
  <tr>
    <td class="ctd" width="15%">定型复改</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复改开始</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复改结束</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型审核</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核开始</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型审核结束</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM开始</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM结束</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <%case "全套复查"%>
  <tr>
    <td class="ctd" width="15%">模头复查</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复查开始</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复查结束</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头BOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM开始</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM结束</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型复查</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复查开始</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复查结束</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM开始</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM结束</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <% If (not isnull(rs("gjshr"))) Then%>
  <tr>
    <td class="ctd" width="15%">共挤复查</td>
    <td class="ctd" width="9%"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">共挤复查开始</td>
    <td class="ctd" width="20%"><%=rs("gjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">共挤复查结束</td>
    <td class="ctd" width="20%"><%=rs("gjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <%case "模头复查"%>
  <tr>
    <td class="ctd" width="15%">模头复查</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复查开始</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头复查结束</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头BOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM开始</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头BOM结束</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <%case "定型复查"%>
  <tr>
    <td class="ctd" width="15%">定型复查</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复查开始</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型复查结束</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM开始</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型BOM结束</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <%end select%>
</table>
<%dim strgy
strgy=""
Response.Write(XjLine(5, "100%", ""))
If not(isNull(rs("gysjr"))) Then
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">工艺设计</td>
    <td class="ctd" width="9%" ><%=rs("gysjr")%>&nbsp;</td>
    <td class="ctd" width="18%">开始时间</td>
    <td class="ctd" width="20%"><%=rs("gysjks")%>&nbsp;</td>
    <td class="ctd" width="18%">结束时间</td>
    <td class="ctd" width="20%"><%=rs("gysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">工艺审核</td>
    <td class="ctd" width="9%" ><%=rs("gyshr")%>&nbsp;</td>
    <td class="ctd" width="18%">开始时间</td>
    <td class="ctd" width="20%"><%=rs("gyshks")%>&nbsp;</td>
    <td class="ctd" width="18%">结束时间</td>
    <td class="ctd" width="20%"><%=rs("gyshjs")%>&nbsp;</td>
  </tr>
</table>
<%
else
select case rs("mjxx")
	case "模头"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">模头工艺设计</td>
    <td class="ctd" width="9%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" >开始时间</td>
    <td class="ctd" width="20%"><%=rs("mtgysjks")%>&nbsp;</td>
    <td class="ctd" >结束时间</td>
    <td class="ctd" width="20%"><%=rs("mtgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">模头工艺审核</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
    <td class="ctd" >结束时间</td>
    <td class="ctd"><%=rs("mtgyshks")%>&nbsp;</td>
    <td class="ctd" >结束时间</td>
    <td class="ctd"><%=rs("mtgyshjs")%>&nbsp;</td>
  </tr>
</table>
<%
	case "定型"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd">定型工艺设计</td>
    <td class="ctd"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgysjks")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">定型工艺审核</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgyshks")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgyshjs")%>&nbsp;</td>
  </tr>
</table>
<%
	case else
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">模头工艺设计</td>
    <td class="ctd" width="9%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" >开始时间</td>
    <td class="ctd" width="20%"><%=rs("mtgysjks")%>&nbsp;</td>
    <td class="ctd" >结束时间</td>
    <td class="ctd" width="20%"><%=rs("mtgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">模头工艺审核</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
    <td class="ctd" >结束时间</td>
    <td class="ctd"><%=rs("mtgyshks")%>&nbsp;</td>
    <td class="ctd" >结束时间</td>
    <td class="ctd"><%=rs("mtgyshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">定型工艺设计</td>
    <td class="ctd"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgysjks")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">定型工艺审核</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgyshks")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("dxgyshjs")%>&nbsp;</td>
  </tr>
  <%If Rs("gjgysjr")<>"" Then%>
  <tr>
    <td class="ctd">共挤工艺设计</td>
    <td class="ctd"><%=rs("gjgysjr")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("gjgysjks")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("gjgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">共挤工艺审核</td>
    <td class="ctd"><%=rs("gjgyshr")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("gjgyshks")%>&nbsp;</td>
    <td class="ctd">结束时间</td>
    <td class="ctd"><%=rs("gjgyshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
</table>
<%
End select
End If
end function

function atask_alluserinfo(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <%select case rs("mjxx")%>
  <%case "全套"%>
  <tr>
    <td class="ctd" width="15%">模头调试单</td>
    <td class="ctd" width="9%"><%=rs("mttsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试单开始</td>
    <td class="ctd" width="20%"><%=rs("mttsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试单结束</td>
    <td class="ctd" width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头调试</td>
    <td class="ctd" width="9%"><%=rs("mttsr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试开始</td>
    <td class="ctd" width="20%"><%=rs("mttsks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试结束</td>
    <td class="ctd" width="20%"><%=rs("mttsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头调试信息整理</td>
    <td class="ctd" width="9%"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试信息整理开始</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试信息整理结束</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型调试单</td>
    <td class="ctd" width="9%"><%=rs("dxtsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试单开始</td>
    <td class="ctd" width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试单结束</td>
    <td class="ctd" width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型调试</td>
    <td class="ctd" width="9%"><%=rs("dxtsr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试开始</td>
    <td class="ctd" width="20%"><%=rs("dxtsks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试结束</td>
    <td class="ctd" width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型调试信息整理</td>
    <td class="ctd" width="9%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试信息整理开始</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试信息整理结束</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">齐套信息整理</td>
    <td class="ctd" width="9%"><%=rs("xtxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">齐套信息整理开始</td>
    <td class="ctd" width="20%"><%=rs("xtxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">齐套信息整理结束</td>
    <td class="ctd" width="20%"><%=rs("xtxxsjjs")%>&nbsp;</td>
  </tr>
  <%case "模头"%>
  <tr>
    <td class="ctd" width="15%">模头调试单</td>
    <td class="ctd" width="9%"><%=rs("mttsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试单开始</td>
    <td class="ctd" width="20%"><%=rs("mttsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试单结束</td>
    <td class="ctd" width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头调试</td>
    <td class="ctd" width="9%"><%=rs("mttsr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试开始</td>
    <td class="ctd" width="20%"><%=rs("mttsks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试结束</td>
    <td class="ctd" width="20%"><%=rs("mttsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">模头调试信息整理</td>
    <td class="ctd" width="9%"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试信息整理开始</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">模头调试信息整理结束</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">齐套信息整理</td>
    <td class="ctd" width="9%"><%=rs("xtxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">齐套信息整理开始</td>
    <td class="ctd" width="20%"><%=rs("xtxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">齐套信息整理结束</td>
    <td class="ctd" width="20%"><%=rs("xtxxsjjs")%>&nbsp;</td>
  </tr>
  <%case "定型"%>
  <tr>
    <td class="ctd" width="15%">定型调试单</td>
    <td class="ctd" width="9%"><%=rs("dxtsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试单开始</td>
    <td class="ctd" width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试单结束</td>
    <td class="ctd" width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型调试</td>
    <td class="ctd" width="9%"><%=rs("dxtsr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试开始</td>
    <td class="ctd" width="20%"><%=rs("dxtsks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试结束</td>
    <td class="ctd" width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">定型调试信息整理</td>
    <td class="ctd" width="9%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试信息整理开始</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">定型调试信息整理结束</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">齐套信息整理</td>
    <td class="ctd" width="9%"><%=rs("xtxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">齐套信息整理开始</td>
    <td class="ctd" width="20%"><%=rs("xtxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">齐套信息整理结束</td>
    <td class="ctd" width="20%"><%=rs("xtxxsjjs")%>&nbsp;</td>
  </tr>
  <%
				case else
				response.write(rs("mjxx"))
			end select
			%>
  <tr>
    <td>
  </tr>
    </td>

    </tr>

</table>
<%
end function

Function DisTd(strfield1,strfield2,strdiff,rs)
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <%if isnull(strfield2) then
  	  if datediff("d", now, rs("jhjssj")) < strdiff then%>
    <td height="7" bgcolor="red"></td>
    <%else
   		 	if isnull(strfield1) then%>
    <td height="7"></td>
    <%else%>
    <td height="7" bgcolor="#6F87CC"></td>
    <%end if
   	 end if
    else
    	if datediff("d", strfield2, rs("jhjssj")) < strdiff then%>
    <td height="7" bgcolor="#991100"></td>
    <%else%>
    <td height="7" bgcolor="#338833"></td>
    <%end if
    end if%>
  </tr>
</table>
<%
End Function

'结构,设计独立计时16:12 2007-4-1-星期日
Function DisTdjg(strfield1,strfield2,strdiff,rs)
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <%if isnull(strfield2) then
  		  if datediff("d", strdiff, now) > 0 then%>
    <td height="7" bgcolor="red"></td>
    <%else
   		 	if isnull(strfield1) then%>
    <td height="7">&nbsp;</td>
    <%else%>
    <td height="7" bgcolor="#6F87CC"></td>
    <%end if
   		 end if
 	 else
		if datediff("d", strdiff, strfield2) > 0 then%>
    <td height="7" bgcolor="#991100"></td>
    <%else%>
    <td height="7" bgcolor="#338833"></td>
    <%end if
  	end if%>
  </tr>
</table>
<%
End Function

Function DisTd2(strfield1,strfield2,rs)
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <%if isnull(strfield2) then
    	if isnull(strfield1) then%>
    <td height="7"></td>
    <%else%>
    <td height="7" bgcolor="#6F87CC"></td>
    <%end if
    else%>
    <td height="7" bgcolor="#338833"></td>
    <%end if%>
  </tr>
</table>
<%
End Function

Function CutLine()	'图例
%>
<table width="95%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td align="right"><table cellpadding="2" cellspacing="2" border="0">
        <tr>
          <td align="right">图例:</td>
          <td class="ctd" bgcolor="#6F87CC"><font color="white">计划内执行</font></td>
          <td class="ctd" bgcolor="#338833"><font color="white">正常完成</font></td>
          <td class="ctd" bgcolor="#991100"><font color="white">过期完成</font></td>
          <td class="ctd" bgcolor="#ff0000"><font color="white">过期未完成</font></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
End Function

Function DisFzInfo(Rs)	'显示分值信息
	Dim Dismtfz, Disdxfz, Disgjfz, DisTslb, DisTsxx, DisTssx, DisSql, mtgjf, dxgjf, DisRs, ssgjf, qbfgjf, qgjf, hgjf
	Dismtfz="" : Disdxfz="" : Disgjfz="" : DisTslb="" : DisTsxx=0 : DisTssx=0 : mtgjf=0 : dxgjf=0
	ssgjf=NullToNum(Rs("ssgj"))
	qbfgjf=NullToNum(Rs("qbfgj"))
	qgjf=NullToNum(Rs("qgj"))
	hgjf=NullToNum(Rs("hgj"))

	select case ssgjf&qbfgjf&qgjf&hgjf
		Case "0000"			'兼容08版共挤计分模式
			'只有软硬前共挤的分值才部分加到模头部分加到定型上
			if Rs("gjfs")="3" and Rs("qhgj")="1" Then
				Dismtfz=Rs("mjzf")*Rs("mtbl")/100
				Disdxfz=Rs("mjzf")*(100-Rs("mtbl"))/100
			End if
			'软硬后共挤的分值单独加到后共挤人上
			If Rs("gjfs")="3" and Rs("qhgj")="2" Then
				Dismtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100
				Disdxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
				Disgjfz=Rs("gjzf")
			End if
			'其他情况下如果有共挤则分全加到模头
			If (not (Rs("gjfs")="3")) Then
				Dismtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + Rs("gjzf")
				Disdxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
		Case Else		'09版共挤计分模式
			If qgjf<>0 Then
				mtgjf=qgjf*Rs("mtbl")/100
				dxgjf=qgjf-mtgjf
			End If
			mtgjf=mtgjf+ssgjf+qbfgjf
			Dismtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + mtgjf
			Disdxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100 + dxgjf
			Disgjfz=Rs("hgj")
	end select
	'调试类别
	DisTslb=Rs("TSLB")
	If ((rs("gjfs")=3) and (rs("qhgj")=2)) or NullToNum(Rs("hgj"))<>0  Then
		DisFzInfo="模具总分: <b>" & Rs("mjzf") & "</b> 分<br>" &_
		"模头分值: <b>" & Dismtfz & "</b> 分<br>" &_
		"定型分值: <b>" & Disdxfz & "</b> 分<br>" &_
		"后共挤分值: <b>" & Disgjfz & "</b> 分<br>" &_
		"BOM分值: <b>" & Rs("bomzf") & "</b> 分<br>" &_
		"调试单分值: <b>" & Rs("tsdzf") & "</b> 分<br>" &_
		"调试分值: <b>" & Rs("tszf") & "</b> 分<br>" &_
		"调试信息整理分值: <b>" & Rs("tsxxzlzf") & "</b> 分"
	Else
	DisFzInfo="模具总分: <b>" & Rs("mjzf") & "</b> 分<br>" &_
		"模头分值: <b>" & Dismtfz & "</b> 分<br>" &_
		"定型分值: <b>" & Disdxfz & "</b> 分<br>" &_
		"BOM分值: <b>" & Rs("bomzf") & "</b> 分<br>" &_
		"调试单分值: <b>" & Rs("tsdzf") & "</b> 分<br>" &_
		"调试分值: <b>" & Rs("tszf") & "</b> 分<br>" &_
		"调试信息整理分值: <b>" & Rs("tsxxzlzf") & "</b> 分"
	End if
	If not(isNull(DisTslb)) Then
		DisSql="select * from [c_tscs] where dmlb like '%"&DisTslb&"%'"
		Set DisRs=xjweb.Exec(DisSql, 1)
		If not DisRs.eof Then
			DisTsxx=DisRs("edxx")
			DisTssx=DisRs("edsx")
		End If
		DisRs.Close
		set DisRs = nothing
		set DisSql = nothing
		DisFzInfo=DisFzInfo &"<br>额定调试次数: <b>"& DisTsxx &" - "& DisTssx &"</b>次"
	End If
End Function
%>
