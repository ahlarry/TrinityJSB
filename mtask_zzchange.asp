<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(4)
CurPage="设计任务 → 更改责任人"
strPage="mtask"
Call FileInc(0, "js/mtask.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
Call closeObj()

Sub Main()
%>
<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchLsh()%></td>
  </tr>
  <Tr>
    <Td class=ctd height=300><%Call mtaskZzchange()%>
      <%Response.Write(XjLine(5,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function mtaskZzchange()
	Dim s_lsh
	s_lsh=""
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("请输入要更改的任务书的流水号!") : Exit Function
	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call closeObj()
		Call JsAlert("流水号为 【" & s_lsh & "】 的任务书不存在!", "mtask_zzchange.asp")
	Else
		If Not IsNull(Rs("sjjssj")) Then
			Call closeObj()
			Call JsAlert("流水号为 【" & s_lsh & "】 的任务书已经完成!", "mtask_zzchange.asp")
		Else
			If Rs("group")=session("userGroup") Or Rs("zz")=Session("userName") or Rs("jgzz")=Session("userName") or Rs("sjzz")=Session("userName") then
				Call mtask_zzchange(rs)
			Else
				Call JsAlert("流水号为 【" & s_lsh & "】 的任务书是 " & Rs("jgzz") &  "、" & Rs("sjzz") & "组的任务!\nSorry! 您无权更改!", "mtask_zzchange.asp")
			End If
		End If
	End If
End Function

Function mtask_zzchange(rs)
	Call mtask_fewinfo(rs)
	Response.Write(XjLine(5,"100%",""))
	Response.Write(XjLine(1,"100%",web_info(12)))
%>
<%Call TbTopic("更改流水号 <font style=color:#0000FF>" &rs("lsh") & "</font> 的任务书") %>
<%If ChkAble(3) Then Response.Write "<a href=mtask_change.asp?s_lsh="&rs("lsh")&">经理权限</a><br>"%>
<table class=xtable cellspacing=0 cellpadding=3 width="95%">
  <form action="mtask_zzchangeindb.asp" method=post>
    <tr>
      <%
	select case rs("mjxx") & rs("rwlr")
			case "全套设计"
	%>
      <td class=ctd>模头结构</td>
      <td class=ctd width="15%"><%if isnull(rs("mtjgr")) then%>
        &nbsp;
        <%else%>
        <select name=mtjgr>
          <option value=<%=rs("mtjgr")%>><%=rs("mtjgr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtjgsj <%if isnull(rs("mtjgjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtjgjs"))) then%>
          <option value=<%=Replace(rs("mtjgjs")," ",".")%>><%=rs("mtjgjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd>模头设计</td>
      <td class=ctd width="15%"><%if isnull(rs("mtsjr")) then%>
        &nbsp;
        <%else%>
        <select name=mtsjr>
          <option value=<%=rs("mtsjr")%>><%=rs("mtsjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtsjsj <%if isnull(rs("mtsjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtsjjs"))) then%>
          <option value=<%=Replace(rs("mtsjjs")," ",".")%>><%=rs("mtsjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <%if not isnull(rs("mtshr")) then%>
      <td class=ctd> 模头审核 </td>
      <td class=ctd width="15%"><select name=mtshr>
          <option value=<%=rs("mtshr")%>><%=rs("mtshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtshsj <%if isnull(rs("mtshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtshjs"))) then%>
          <option value=<%=Replace(rs("mtshjs")," ",".")%>><%=rs("mtshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <%end if%>
      <td class=ctd> 模头BOM </td>
      <td class=ctd width="15%"><%if isnull(rs("mtbomr")) then%>
        &nbsp;
        <%else%>
        <select name=mtbomr>
          <option value=<%=rs("mtbomr")%>><%=rs("mtbomr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtbomsj <%if isnull(rs("mtbomjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtbomjs"))) then%>
          <option value=<%=Replace(rs("mtbomjs")," ",".")%>><%=rs("mtbomjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
    </tr>
    <%if not isnull(rs("mtjgshr")) then%>
    <tr>
      <td class=ctd> 模头结构确认 </td>
      <td class=ctd width="15%"><select name=mtjgshr>
          <option value=<%=rs("mtjgshr")%>><%=rs("mtjgshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtjgshsj <%if isnull(rs("mtjgshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtjgshjs"))) then%>
          <option value=<%=Replace(rs("mtjgshjs")," ",".")%>><%=rs("mtjgshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <td class=ctd> 模头设计确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("mtsjshr")) then%>
        &nbsp;
        <%else%>
        <select name=mtsjshr>
          <option value=<%=rs("mtsjshr")%>><%=rs("mtsjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtsjshsj <%if isnull(rs("mtsjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtsjshjs"))) then%>
          <option value=<%=Replace(rs("mtsjshjs")," ",".")%>><%=rs("mtsjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%end if%>
    <tr>
      <td class=ctd> 定型结构 </td>
      <td class=ctd width="15%"><%if isnull(rs("dxjgr")) then%>
        &nbsp;
        <%else%>
        <select name=dxjgr>
          <option value=<%=rs("dxjgr")%>><%=rs("dxjgr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxjgsj <%if isnull(rs("dxjgjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxjgjs"))) then%>
          <option value=<%=Replace(rs("dxjgjs")," ",".")%>><%=rs("dxjgjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd> 定型设计 </td>
      <td class=ctd width="15%"><%if isnull(rs("dxsjr")) then%>
        &nbsp;
        <%else%>
        <select name=dxsjr>
          <option value=<%=rs("dxsjr")%>><%=rs("dxsjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxsjsj <%if isnull(rs("dxsjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxsjjs"))) then%>
          <option value=<%=Replace(rs("dxsjjs")," ",".")%>><%=rs("dxsjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <%if not isnull(rs("dxshr")) then%>
      <td class=ctd> 定型审核 </td>
      <td class=ctd width="15%"><select name=dxshr>
          <option value=<%=rs("dxshr")%>><%=rs("dxshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxshsj <%if isnull(rs("dxshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxshjs"))) then%>
          <option value=<%=Replace(rs("dxshjs")," ",".")%>><%=rs("dxshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <%end if%>
      <td class=ctd width="10%"> 定型BOM </td>
      <td class=ctd width="15%"><%if isnull(rs("dxbomr")) then%>
        &nbsp;
        <%else%>
        <select name=dxbomr>
          <option value=<%=rs("dxbomr")%>><%=rs("dxbomr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxbomsj <%if isnull(rs("dxbomjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxbomjs"))) then%>
          <option value=<%=Replace(rs("dxbomjs")," ",".")%>><%=rs("dxbomjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
    </tr>
    <%if not isnull(rs("dxjgshr")) then%>
    <tr>
      <td class=ctd> 定型结构确认 </td>
      <td class=ctd width="15%"><select name=dxjgshr>
          <option value=<%=rs("dxjgshr")%>><%=rs("dxjgshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxjgshsj <%if isnull(rs("dxjgshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxjgshjs"))) then%>
          <option value=<%=Replace(rs("dxjgshjs")," ",".")%>><%=rs("dxjgshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <td class=ctd> 定型设计确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("dxsjshr")) then%>
        &nbsp;
        <%else%>
        <select name=dxsjshr>
          <option value=<%=rs("dxsjshr")%>><%=rs("dxsjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxsjshsj <%if isnull(rs("dxsjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxsjshjs"))) then%>
          <option value=<%=Replace(rs("dxsjshjs")," ",".")%>><%=rs("dxsjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%end if%>
    <% If ((not isnull(rs("gjjgr"))) or (not isnull(rs("gjsjr")))) Then %>
    <tr>
      <td class=ctd> 后共挤结构 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjjgr")) then%>
        &nbsp;
        <%else%>
        <select name=gjjgr>
          <option value=<%=rs("gjjgr")%>><%=rs("gjjgr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjjgsj <%if isnull(rs("gjjgjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjjgjs"))) then%>
          <option value=<%=Replace(rs("gjjgjs")," ",".")%>><%=rs("gjjgjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd width="10%"> 后共挤设计 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjsjr")) then%>
        &nbsp;
        <%else%>
        <select name=gjsjr>
          <option value=<%=rs("gjsjr")%>><%=rs("gjsjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjsjsj <%if isnull(rs("gjsjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjsjjs"))) then%>
          <option value=<%=Replace(rs("gjsjjs")," ",".")%>><%=rs("gjsjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <%if isnull(rs("gjshr")) then%>
      <td class=ctd colspan=2>&nbsp;</td>
      <%else%>
      <td class=ctd> 后共挤审核 </td>
      <td class=ctd width="15%"><select name=gjshr>
          <option value=<%=rs("gjshr")%>><%=rs("gjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjshsj <%if isnull(rs("gjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjshjs"))) then%>
          <option value=<%=Replace(rs("gjshjs")," ",".")%>><%=rs("gjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <%end if%>
    </tr>
    <%End if%>
    <%if not isnull(rs("gjjgshr")) then%>
    <tr>
      <td class=ctd> 后共挤结构确认 </td>
      <td class=ctd width="15%"><select name=gjjgshr>
          <option value=<%=rs("gjjgshr")%>><%=rs("gjjgshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjjgshsj <%if isnull(rs("gjjgshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjjgshjs"))) then%>
          <option value=<%=Replace(rs("gjjgshjs")," ",".")%>><%=rs("gjjgshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <td class=ctd> 后共挤设计确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjsjshr")) then%>
        &nbsp;
        <%else%>
        <select name=gjsjshr>
          <option value=<%=rs("gjsjshr")%>><%=rs("gjsjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjsjshsj <%if isnull(rs("gjsjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjsjshjs"))) then%>
          <option value=<%=Replace(rs("gjsjshjs")," ",".")%>><%=rs("gjsjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%end if
				case "全套复改"
				%>
    <td class=ctd width="15%" rowspan=2>&nbsp;</td>
    <td class=ctd width="15%" rowspan=2>&nbsp;</td>
    <td class=ctd width="10%"> 模头复改 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtsjr")) then%>
      &nbsp;
      <%else%>
      <select name=mtsjr>
        <option value=<%=rs("mtsjr")%>><%=rs("mtsjr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtsjsj <%if isnull(rs("mtsjjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtsjjs"))) then%>
        <option value=<%=Replace(rs("mtsjjs")," ",".")%>><%=rs("mtsjjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头审核 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtshr")) then%>
      &nbsp;
      <%else%>
      <select name=mtshr>
        <option value=<%=rs("mtshr")%>><%=rs("mtshr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtshsj <%if isnull(rs("mtshjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtshjs"))) then%>
        <option value=<%=Replace(rs("mtshjs")," ",".")%>><%=rs("mtshjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("mtbomr")) then%>
      &nbsp;
      <%else%>
      <select name=mtbomr>
        <option value=<%=rs("mtbomr")%>><%=rs("mtbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtbomsj <%if isnull(rs("mtbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtbomjs"))) then%>
        <option value=<%=Replace(rs("mtbomjs")," ",".")%>><%=rs("mtbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型复改 </td>
    <td class=ctd width="15%"><%if isnull(rs("dxsjr")) then%>
      &nbsp;
      <%else%>
      <select name=dxsjr>
        <option value=<%=rs("dxsjr")%>><%=rs("dxsjr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxsjsj <%if isnull(rs("dxsjjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxsjjs"))) then%>
        <option value=<%=Replace(rs("dxsjjs")," ",".")%>><%=rs("dxsjjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型审核 </td>
    <td class=ctd width="15%"><%if isnull(rs("dxshr")) then%>
      &nbsp;
      <%else%>
      <select name=dxshr>
        <option value=<%=rs("dxshr")%>><%=rs("dxshr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxshsj <%if isnull(rs("dxshjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxshjs"))) then%>
        <option value=<%=Replace(rs("dxshjs")," ",".")%>><%=rs("dxshjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("dxbomr")) then%>
      &nbsp;
      <%else%>
      <select name=dxbomr>
        <option value=<%=rs("dxbomr")%>><%=rs("dxbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxbomsj <%if isnull(rs("dxbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxbomjs"))) then%>
        <option value=<%=Replace(rs("dxbomjs")," ",".")%>><%=rs("dxbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <%
				case "全套复查"
				%>
    <td class=ctd width="10%"  colspan=4 rowspan=2>&nbsp;</td>
    <td class=ctd width="10%"> 模头复查 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtshr")) then%>
      &nbsp;
      <%else%>
      <select name=mtshr>
        <option value=<%=rs("mtshr")%>><%=rs("mtshr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtshsj <%if isnull(rs("mtshjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtshjs"))) then%>
        <option value=<%=Replace(rs("mtshjs")," ",".")%>><%=rs("mtshjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("mtbomr")) then%>
      &nbsp;
      <%else%>
      <select name=mtbomr>
        <option value=<%=rs("mtbomr")%>><%=rs("mtbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtbomsj <%if isnull(rs("mtbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtbomjs"))) then%>
        <option value=<%=Replace(rs("mtbomjs")," ",".")%>><%=rs("mtbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
      </tr>

    <tr>
      <td class=ctd width="10%"> 定型复查 </td>
      <td class=ctd width="15%"><%if isnull(rs("dxshr")) then%>
        &nbsp;
        <%else%>
        <select name=dxshr>
          <option value=<%=rs("dxshr")%>><%=rs("dxshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxshsj <%if isnull(rs("dxshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxshjs"))) then%>
          <option value=<%=Replace(rs("dxshjs")," ",".")%>><%=rs("dxshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd width="10%"> 定型BOM </td>
      <td class=ctd width="15%"><%if isnull(rs("dxbomr")) then%>
        &nbsp;
        <%else%>
        <select name=dxbomr>
          <option value=<%=rs("dxbomr")%>><%=rs("dxbomr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxbomsj <%if isnull(rs("dxbomjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxbomjs"))) then%>
          <option value=<%=Replace(rs("dxbomjs")," ",".")%>><%=rs("dxbomjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
    </tr>
    <%if not isnull(rs("gjshr")) then%>
    <tr>
      <td class=ctd colspan=4>&nbsp;</td>
      <td class=ctd> 共挤复查 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjshr")) then%>
        &nbsp;
        <%else%>
        <select name=gjfcr>
          <option value=<%=rs("gjshr")%>><%=rs("gjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjfcjs <%if isnull(rs("gjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjshjs"))) then%>
          <option value=<%=Replace(rs("gjshjs")," ",".")%>><%=rs("gjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%end if
				case "模头设计"
				%>
    <td class=ctd width="10%"> 模头结构 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtjgr")) then%>
      &nbsp;
      <%else%>
      <select name=mtjgr>
        <option value=<%=rs("mtjgr")%>><%=rs("mtjgr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtjgsj <%if isnull(rs("mtjgjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtjgjs"))) then%>
        <option value=<%=Replace(rs("mtjgjs")," ",".")%>><%=rs("mtjgjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头设计 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtsjr")) then%>
      &nbsp;
      <%else%>
      <select name=mtsjr>
        <option value=<%=rs("mtsjr")%>><%=rs("mtsjr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtsjsj <%if isnull(rs("mtsjjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtsjjs"))) then%>
        <option value=<%=Replace(rs("mtsjjs")," ",".")%>><%=rs("mtsjjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("mtbomr")) then%>
      &nbsp;
      <%else%>
      <select name=mtbomr>
        <option value=<%=rs("mtbomr")%>><%=rs("mtbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtbomsj <%if isnull(rs("mtbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtbomjs"))) then%>
        <option value=<%=Replace(rs("mtbomjs")," ",".")%>><%=rs("mtbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
      </tr>

    <tr>
      <td class=ctd> 模头结构确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("mtjgshr")) then%>
        &nbsp;
        <%else%>
        <select name=mtjgshr>
          <option value=<%=rs("mtjgshr")%>><%=rs("mtjgshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtjgshsj <%if isnull(rs("mtjgshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtjgshjs"))) then%>
          <option value=<%=Replace(rs("mtjgshjs")," ",".")%>><%=rs("mtjgshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%End If%></td>
      <td class=ctd> 模头设计确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("mtsjshr")) then%>
        &nbsp;
        <%else%>
        <select name=mtsjshr>
          <option value=<%=rs("mtsjshr")%>><%=rs("mtsjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtsjshsj <%if isnull(rs("mtsjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtsjshjs"))) then%>
          <option value=<%=Replace(rs("mtsjshjs")," ",".")%>><%=rs("mtsjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <% If ((not isnull(rs("gjjgr"))) or (not isnull(rs("gjsjr")))) Then %>
    <tr>
      <td class=ctd> 后共挤结构 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjjgr")) then%>
        &nbsp;
        <%else%>
        <select name=gjjgr>
          <option value=<%=rs("gjjgr")%>><%=rs("gjjgr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjjgsj <%if isnull(rs("gjjgjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjjgjs"))) then%>
          <option value=<%=Replace(rs("gjjgjs")," ",".")%>><%=rs("gjjgjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd width="10%"> 后共挤设计 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjsjr")) then%>
        &nbsp;
        <%else%>
        <select name=gjsjr>
          <option value=<%=rs("gjsjr")%>><%=rs("gjsjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjsjsj <%if isnull(rs("gjsjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjsjjs"))) then%>
          <option value=<%=Replace(rs("gjsjjs")," ",".")%>><%=rs("gjsjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <%if isnull(rs("gjshr")) then%>
      <td class=ctd colspan=2>&nbsp;</td>
      <%else%>
      <td class=ctd> 后共挤审核 </td>
      <td class=ctd width="15%"><select name=gjshr>
          <option value=<%=rs("gjshr")%>><%=rs("gjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjshsj <%if isnull(rs("gjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjshjs"))) then%>
          <option value=<%=Replace(rs("gjshjs")," ",".")%>><%=rs("gjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <%end if%>
    </tr>
    <%End if%>
    <%if not isnull(rs("gjjgshr")) then%>
    <tr>
      <td class=ctd> 后共挤结构确认 </td>
      <td class=ctd width="15%"><select name=gjjgshr>
          <option value=<%=rs("gjjgshr")%>><%=rs("gjjgshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjjgshsj <%if isnull(rs("gjjgshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjjgshjs"))) then%>
          <option value=<%=Replace(rs("gjjgshjs")," ",".")%>><%=rs("gjjgshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <td class=ctd> 后共挤设计确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjsjshr")) then%>
        &nbsp;
        <%else%>
        <select name=gjsjshr>
          <option value=<%=rs("gjsjshr")%>><%=rs("gjsjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjsjshsj <%if isnull(rs("gjsjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjsjshjs"))) then%>
          <option value=<%=Replace(rs("gjsjshjs")," ",".")%>><%=rs("gjsjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%end if
				case "模头复改"
				%>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="15%">&nbsp;</td>
    ?
    <td class=ctd width="10%"> 模头复改 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtsjr")) then%>
      &nbsp;
      <%else%>
      <select name=mtsjr>
        <option value=<%=rs("mtsjr")%>><%=rs("mtsjr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtsjsj <%if isnull(rs("mtsjjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtsjjs"))) then%>
        <option value=<%=Replace(rs("mtsjjs")," ",".")%>><%=rs("mtsjjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头审核 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtshr")) then%>
      &nbsp;
      <%else%>
      <select name=mtshr>
        <option value=<%=rs("mtshr")%>><%=rs("mtshr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtshsj <%if isnull(rs("mtshjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtshjs"))) then%>
        <option value=<%=Replace(rs("mtshjs")," ",".")%>><%=rs("mtshjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("mtbomr")) then%>
      &nbsp;
      <%else%>
      <select name=mtbomr>
        <option value=<%=rs("mtbomr")%>><%=rs("mtbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtbomsj <%if isnull(rs("mtbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtbomjs"))) then%>
        <option value=<%=Replace(rs("mtbomjs")," ",".")%>><%=rs("mtbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <%
				case "模头复查"
				%>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="15%">&nbsp;</td>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="15%">&nbsp;</td>
    <td class=ctd width="10%"> 模头复查 </td>
    <td class=ctd width="15%"><%if isnull(rs("mtshr")) then%>
      &nbsp;
      <%else%>
      <select name=mtshr>
        <option value=<%=rs("mtshr")%>><%=rs("mtshr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtshsj <%if isnull(rs("mtshjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtshjs"))) then%>
        <option value=<%=Replace(rs("mtshjs")," ",".")%>><%=rs("mtshjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 模头BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("mtbomr")) then%>
      &nbsp;
      <%else%>
      <select name=mtbomr>
        <option value=<%=rs("mtbomr")%>><%=rs("mtbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=mtbomsj <%if isnull(rs("mtbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("mtbomjs"))) then%>
        <option value=<%=Replace(rs("mtbomjs")," ",".")%>><%=rs("mtbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <%
				case "定型设计"
				%>
    <td class=ctd width="10%"> 定型结构 </td>
    <td class=ctd width="15%"><%if isnull(rs("dxjgr")) then%>
      &nbsp;
      <%else%>
      <select name=dxjgr>
        <option value=<%=rs("dxjgr")%>><%=rs("dxjgr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxjgsj <%if isnull(rs("dxjgjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxjgjs"))) then%>
        <option value=<%=Replace(rs("dxjgjs")," ",".")%>><%=rs("dxjgjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型设计 </td>
    <td class=ctd width="15%"><%if isnull(rs("dxsjr")) then%>
      &nbsp;
      <%else%>
      <select name=dxsjr>
        <option value=<%=rs("dxsjr")%>><%=rs("dxsjr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxsjsj <%if isnull(rs("dxsjjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxsjjs"))) then%>
        <option value=<%=Replace(rs("dxsjjs")," ",".")%>><%=rs("dxsjjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("dxbomr")) then%>
      &nbsp;
      <%else%>
      <select name=dxbomr>
        <option value=<%=rs("dxbomr")%>><%=rs("dxbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxbomsj <%if isnull(rs("dxbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxbomjs"))) then%>
        <option value=<%=Replace(rs("dxbomjs")," ",".")%>><%=rs("dxbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
      </tr>

    <tr>
      <td class=ctd> 定型结构确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("dxjgshr")) then%>
        &nbsp;
        <%else%>
        <select name=dxjgshr>
          <option value=<%=rs("dxjgshr")%>><%=rs("dxjgshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxjgshsj <%if isnull(rs("dxjgshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxjgshjs"))) then%>
          <option value=<%=Replace(rs("dxjgshjs")," ",".")%>><%=rs("dxjgshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%End If%></td>
      <td class=ctd> 定型设计确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("dxsjshr")) then%>
        &nbsp;
        <%else%>
        <select name=dxsjshr>
          <option value=<%=rs("dxsjshr")%>><%=rs("dxsjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxsjshsj <%if isnull(rs("dxsjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxsjshjs"))) then%>
          <option value=<%=Replace(rs("dxsjshjs")," ",".")%>><%=rs("dxsjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
        <% If ((not isnull(rs("gjjgr"))) or (not isnull(rs("gjsjr")))) Then %>
    <tr>
      <td class=ctd> 后共挤结构 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjjgr")) then%>
        &nbsp;
        <%else%>
        <select name=gjjgr>
          <option value=<%=rs("gjjgr")%>><%=rs("gjjgr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjjgsj <%if isnull(rs("gjjgjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjjgjs"))) then%>
          <option value=<%=Replace(rs("gjjgjs")," ",".")%>><%=rs("gjjgjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd width="10%"> 后共挤设计 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjsjr")) then%>
        &nbsp;
        <%else%>
        <select name=gjsjr>
          <option value=<%=rs("gjsjr")%>><%=rs("gjsjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjsjsj <%if isnull(rs("gjsjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjsjjs"))) then%>
          <option value=<%=Replace(rs("gjsjjs")," ",".")%>><%=rs("gjsjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <%if isnull(rs("gjshr")) then%>
      <td class=ctd colspan=2>&nbsp;</td>
      <%else%>
      <td class=ctd> 后共挤审核 </td>
      <td class=ctd width="15%"><select name=gjshr>
          <option value=<%=rs("gjshr")%>><%=rs("gjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjshsj <%if isnull(rs("gjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjshjs"))) then%>
          <option value=<%=Replace(rs("gjshjs")," ",".")%>><%=rs("gjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <%end if%>
    </tr>
    <%End if%>
    <%if not isnull(rs("gjjgshr")) then%>
    <tr>
      <td class=ctd> 后共挤结构确认 </td>
      <td class=ctd width="15%"><select name=gjjgshr>
          <option value=<%=rs("gjjgshr")%>><%=rs("gjjgshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjjgshsj <%if isnull(rs("gjjgshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjjgshjs"))) then%>
          <option value=<%=Replace(rs("gjjgshjs")," ",".")%>><%=rs("gjjgshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select></td>
      <td class=ctd> 后共挤设计确认 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjsjshr")) then%>
        &nbsp;
        <%else%>
        <select name=gjsjshr>
          <option value=<%=rs("gjsjshr")%>><%=rs("gjsjshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjsjshsj <%if isnull(rs("gjsjshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjsjshjs"))) then%>
          <option value=<%=Replace(rs("gjsjshjs")," ",".")%>><%=rs("gjsjshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%end if%></td>
      <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%end if
				case "定型复改"
				%>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="15%">&nbsp;</td>
    <td class=ctd width="10%"> 定型复改 </td>
    <td class=ctd width="15%"><%if isnull(rs("dxsjr")) then%>
      &nbsp;
      <%else%>
      <select name=dxsjr>
        <option value=<%=rs("dxsjr")%>><%=rs("dxsjr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxsjsj <%if isnull(rs("dxsjjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxsjjs"))) then%>
        <option value=<%=Replace(rs("dxsjjs")," ",".")%>><%=rs("dxsjjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型审核 </td>
    <td class=ctd width="15%"><%if isnull(rs("dxshr")) then%>
      &nbsp;
      <%else%>
      <select name=dxshr>
        <option value=<%=rs("dxshr")%>><%=rs("dxshr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxshsj <%if isnull(rs("dxshjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxshjs"))) then%>
        <option value=<%=Replace(rs("dxshjs")," ",".")%>><%=rs("dxshjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("dxbomr")) then%>
      &nbsp;
      <%else%>
      <select name=dxbomr>
        <option value=<%=rs("dxbomr")%>><%=rs("dxbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxbomsj <%if isnull(rs("dxbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxbomjs"))) then%>
        <option value=<%=Replace(rs("dxbomjs")," ",".")%>><%=rs("dxbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <%
				case "定型复查"
				%>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="15%">&nbsp;</td>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="15%">&nbsp;</td>
    <td class=ctd width="10%"> 定型复查 </td>
    <td class=ctd width="15%"><%if isnull(rs("dxshr")) then%>
      &nbsp;
      <%else%>
      <select name=dxshr>
        <option value=<%=rs("dxshr")%>><%=rs("dxshr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxshsj <%if isnull(rs("dxshjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxshjs"))) then%>
        <option value=<%=Replace(rs("dxshjs")," ",".")%>><%=rs("dxshjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <td class=ctd width="10%"> 定型BOM </td>
    <td class=ctd width="15%"><%if isnull(rs("dxbomr")) then%>
      &nbsp;
      <%else%>
      <select name=dxbomr>
        <option value=<%=rs("dxbomr")%>><%=rs("dxbomr")%></option>
        <%for i=0 to ubound(c_jsb)%>
        <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
        <%next%>
      </select>
      <select name=dxbomsj <%if isnull(rs("dxbomjs")) then%> disabled="disabled" <%End If%>>
        <%if not(isnull(rs("dxbomjs"))) then%>
        <option value=<%=Replace(rs("dxbomjs")," ",".")%>><%=rs("dxbomjs")%></option>
        <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
        <%End If%>
      </select>
      <%end if%></td>
    <%
				case else
				response.write(rs("mjxx") & rs("rwlr"))
			end select
		%>
      </tr>
    <%
    If not(isnull(rs("mtgysjr"))) Then
    %>
    <tr>
      <td class=ctd> 模头工艺设计 </td>
      <td class=ctd width="15%">
        <select name=mtgysjr>
          <option value=<%=rs("mtgysjr")%>><%=rs("mtgysjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtgysjsj <%if isnull(rs("mtgysjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtgysjjs"))) then%>
          <option value=<%=Replace(rs("mtgysjjs")," ",".")%>><%=rs("mtgysjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
      </td>
      <td class=ctd> 模头工艺审核 </td>
      <td class=ctd width="15%"><%if isnull(rs("mtgyshr")) then%>
        &nbsp;
        <%else%>
        <select name=mtgyshr>
          <option value=<%=rs("mtgyshr")%>><%=rs("mtgyshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=mtgyshsj <%if isnull(rs("mtgyshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("mtgyshjs"))) then%>
          <option value=<%=Replace(rs("mtgyshjs")," ",".")%>><%=rs("mtgyshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%End If%>
       </td>
       <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%End If
    If not(isnull(rs("dxgysjr"))) Then
    %>
    <tr>
      <td class=ctd> 定型工艺设计 </td>
      <td class=ctd width="15%">
        <select name=dxgysjr>
          <option value=<%=rs("dxgysjr")%>><%=rs("dxgysjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxgysjsj <%if isnull(rs("dxgysjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxgysjjs"))) then%>
          <option value=<%=Replace(rs("dxgysjjs")," ",".")%>><%=rs("dxgysjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
      </td>
      <td class=ctd> 定型工艺审核 </td>
      <td class=ctd width="15%"><%if isnull(rs("dxgyshr")) then%>
        &nbsp;
        <%else%>
        <select name=dxgyshr>
          <option value=<%=rs("dxgyshr")%>><%=rs("dxgyshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=dxgyshsj <%if isnull(rs("dxgyshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("dxgyshjs"))) then%>
          <option value=<%=Replace(rs("dxgyshjs")," ",".")%>><%=rs("dxgyshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%End If%>
       </td>
       <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%End If

    If not(isnull(rs("gjgysjr"))) Then
    %>
    <tr>
      <td class=ctd> 共挤工艺设计 </td>
      <td class=ctd width="15%">
        <select name=gjgysjr>
          <option value=<%=rs("gjgysjr")%>><%=rs("gjgysjr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjgysjsj <%if isnull(rs("gjgysjjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjgysjjs"))) then%>
          <option value=<%=Replace(rs("gjgysjjs")," ",".")%>><%=rs("gjgysjjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
      </td>
      <td class=ctd> 共挤工艺审核 </td>
      <td class=ctd width="15%"><%if isnull(rs("gjgyshr")) then%>
        &nbsp;
        <%else%>
        <select name=gjgyshr>
          <option value=<%=rs("gjgyshr")%>><%=rs("gjgyshr")%></option>
          <%for i=0 to ubound(c_jsb)%>
          <option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
          <%next%>
        </select>
        <select name=gjgyshsj <%if isnull(rs("gjgyshjs")) then%> disabled="disabled" <%End If%>>
          <%if not(isnull(rs("gjgyshjs"))) then%>
          <option value=<%=Replace(rs("gjgyshjs")," ",".")%>><%=rs("gjgyshjs")%></option>
          <option value=<%=Replace(now()," ",".")%>><%=now()%></option>
          <%End If%>
        </select>
        <%End If%>
       </td>
       <td class=ctd colspan=2>&nbsp;</td>
    </tr>
    <%End If%>

    <tr>
      <td class=ctd colspan=8><input type=submit value=" 更改 "></td>
    </tr>
    <input type=hidden name=lsh value=<%=rs("lsh")%>>
  </form>
</table>
<%
end function
%>
