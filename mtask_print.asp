<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'15:28 2006-11-2-星期四
Call ChkPageAble(0)
CurPage="设计任务 → 打印任务书"
strPage="mtask"
xjweb.header()
'Call TopTable()
Call Main()
'Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<Table border=0 cellspacing=0 cellpadding=2 width="720">
  <Tr>
    <Td height=300><%Call mtaskDisplay()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function mtaskDisplay()
	Dim s_lsh, action
	s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("请确定打印任务书的流水号!") : Exit Function
	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then Call JsAlert("流水号 【" & s_lsh & "】 任务书不存在! 请重新输入流水号!", "mtask_display.asp") : Exit Function
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
<%Call TbTopic("挤出模具厂挤出模设计任务书")%>
<table class=ktable cellspacing=0 cellpadding=3 width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8"><b>合同信息</b></td>
  </tr>
  <tr>
    <td class="rtd" width="13%">订单号</td>
    <td class="ltd"><%=rs("ddh")%></td>
    <td class="rtd" width="13%">流水号</td>
    <td colspan="2" class="ltd" width="*"><a href="SreacDwg.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="rtd" width="13%">模号</td>
    <td colspan="2" class="ltd" width="*"><%=rs("mh")%></td>
  </tr>
  <tr>
    <td class="rtd">客户名称</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">断面名称</td>
    <td colspan="2" class="ltd"><%=rs("dmmc")%></td>
    <td class="rtd">模具材料</td>
    <td colspan="2" class="ltd"><%=rs("mjcl")%></td>
  </tr>
  <tr>
    <td class="rtd">设备厂家</td>
    <td class="ltd"><%=rs("sbcj")%></td>
    <td class="rtd">水接头数量</td>
    <td colspan="2" class="ltd"><%=rs("sjtsl")%></td>
    <td class="rtd">气接头数量</td>
    <td colspan="2" class="ltd"><%=rs("qjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">挤出机型号</td>
    <td class="ltd"><%=rs("jcjxh")%></td>
    <td class="rtd">挤出方向</td>
    <td colspan="2" class="ltd"><%=rs("jcfx")%></td>
    <td class="rtd">牵引速度</td>
    <td colspan="2" class="ltd"><%=rs("qysd")%> 米/分(m/min)</td>
  </tr>
  <tr>
    <td class="rtd">配加热板</td>
    <td class="ltd"><%if rs("pjrb") then%>
      是
      <%else%>
      否
      <%end if%></td>
    <td class="rtd">加热板信息</td>
    <td colspan="2" class="ltd">相数:<%=rs("jrbxs")%> 材质:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
    <td class="rtd">腔数</td>
    <td colspan="2" class="ltd"><%=rs("qs")%>腔</td>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8"><b>模具信息</b></td>
  </tr>
  <tr>
    <td class="rtd"  width="13%">任务内容</td>
    <td class="ltd"><%=rs("mjxx") & rs("rwlr")%></td>
        <td class="rtd"  width="13%">厂内调试</td>
    <td class="ltd"  colspan="2" ><%if rs("cnts") then%>
      是
      <%else%>
      &nbsp;/
      <%end if%></td>
    <td class="rtd"  width="13%">调试类别</td>
    <% If Rs("cnts") Then%>
    <%If Not(isnull(Rs("tslb"))) Then%>
    <td class="ltd"  colspan="2"><a href="mtest_display.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("tslb")%></a></td>
    <%Else%>
    <td class="ltd" colspan="2" >&nbsp;/</td>
    <%End If%>
    <%Else%>
    <%If Rs("beit") Then%>
    <td class="ltd"  colspan="2">北调</td>
    <%Else%>
    <td class="ltd" colspan="2">&nbsp;/</td>
    <%End If%>
    <%End If%>
  </tr>
  <tr>
    <td class="rtd">模头结构</td>
    <td class="ltd"><%if IsNull(rs("mtjg")) Then
    	Response.Write("&nbsp;")
    else
    	Response.Write(rs("mtjg"))
    End if%></td>
    <td class="rtd">定型结构</td>
    <td class="ltd" colspan="2" ><%if IsNull(rs("dxjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("dxjg"))
    End if%></td>
    <td class="rtd">水箱结构</td>
    <td class="ltd" colspan="2" ><%if IsNull(rs("sxjg")) Then
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
    <td class="ltd" colspan="2" ><%=rs("mtljcc")%></td>
    <td class="rtd">热电偶规格</td>
    <td class="ltd" colspan="2" ><%=rs("rdogg")%></td>
  </tr>
  <tr>
    <td class="rtd">共挤类型</td>
    <td class="ltd"><%=Trim(gjxx)%></td>
    <td class="rtd">共挤连接尺寸</td>
    <td class="ltd" colspan="2" ><%=rs("gjljcc")%>&nbsp;</td>
    <td class="rtd">型材壁厚</td>
    <td class="ltd" colspan="2" ><%=Rs("xcbh")%>毫米</td>
  </tr>
  <tr>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8" alt="<%=DisFzInfo(Rs)%>"><b>其他信息</b></td>
  </tr>
  <tr>
    <td class="rtd" height="180">评审记录</td>
    <td class="ltd"colspan="7"><table width="100%" >
        <tr>
          <td><label>
              <input type="checkbox" name="psyx" value="3" id="psyx_2" />
              需评审，结果如下:</label></td>
          <td><label>
              <input type="checkbox" name="psyx" value="1" id="psyx_0" />
              按照设计规范进行设计</label></td>
          <td><label>
              <input type="checkbox" name="psyx" value="2" id="psyx_1" />
              按照客户来图进行设计 </label></td>
        </tr>
        <tr>
          <td height="120" colspan="5" valign="middle"><%=xjweb.HtmlToCode(rs("psjl"))%></Td>
        </tr>
        <tr valign="bottom">
          <td></Td>
          <td colspan="4" valign="bottom"><p align="left">签名:
            </p></Td>
        </tr>
      </table></Td>
  </tr>
  <tr>
    <td class="rtd">备注</td>
    <td class="ltd" colspan="7" height="180" valign="top"><%=xjweb.HtmlToCode(rs("bz"))%></td>
  </tr>
  <tr>
    <td class="rtd">计划开始</td>
    <%If rs("jhkssj")<>"" Then%>
    <td class="ltd"><%=XjDate(rs("jhkssj"),3)%></td>
    <%else%>
    <td class="ltd" width="120">&nbsp;/&nbsp;</td>
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
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call mtask_userinfo(rs)%>
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call atask_userinfo(rs)%>
<%
	Rs.Close
End Function
%>
