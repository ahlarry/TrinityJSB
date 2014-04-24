<!--#include file="include/conn.asp"-->
<!--#include file="js/jsCookie.js"-->
<%
Call ChkPageAble(7)
CurPage="图档备份 → 添加存档"
strPage="docbak"
Call FileInc(0, "js/docbak.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchLsh()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call DocBak()%>
			<%Response.Write(XjLine(10,"100%",""))%>
			<%Response.Write(XjLine(1,"100%",web_info(12)))%>
			<%Call waitSave()%>

			<%Response.Write(XjLine(10,"100%",""))%>

		</Td></Tr>
	</Table>
<%
End Sub

Function DocBak()
	Dim s_lsh
	s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("请选择您要存盘的任务的流水号!") : Exit Function
	strSql="select mh,lsh,ddh,dwmc,dmmc,cp from [mtask] where lsh='"&s_lsh&"' or ddh='"&s_lsh&"' and not(cp)"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("流水号或订单号为 【"&s_lsh&"】 模具任务不存在！","docbak_add.asp")
	ElseIf Rs("cp") Then
		Call JsAlert("流水号或订单号为 【"&s_lsh&"】 模具图纸已经存盘！","docbak_add.asp")
	Else
		Call SelectSave (rs,s_lsh)
	End IF
	Rs.Close
End Function

Function DocBakAdd(rs)
%>

	<%Call TbTopic("添加流水号" & rs("lsh") & " 模具图纸存档信息")%>
	<table width="68%" cellpadding=2 cellspacing=0 class=xtable>
		<form name="frm_docbak" id="frm_docbak" action="docbak_indb.asp" method="post" onsubmit='return docbak_checkinf();'>
		<tr>
			<td class=rtd>模号</td>
			<td class=ltd><%=UCase(rs("mh") & "-" & rs("lsh"))%></td>
		</tr>
		<tr>
			<td class=rtd>单位名称</td>
			<td class=ltd><%=rs("dwmc")%></td>
		</tr>
		<tr>
			<td class=rtd>断面名称</td>
			<td class=ltd><%=rs("dmmc")%></td>
		</tr>
		<tr>
			<td class=rtd>所存盘号</td>
			<td class=ltd><input type="text" name="diskid" size="15"></td>
		</tr>
		<tr>
			<td class=rtd>备注</td>
			<td class=ltd><textarea name="bz" cols="60" rows="10"></textarea></td>
		</tr>
		<tr>
			<td class=ctd colspan="2">
				<input type="hidden" name="lsh1" value="<%=rs("lsh")%>">
				<input type="hidden" name="ddh1" value="<%=rs("ddh")%>">
				<input type="hidden" name="mh1" value="<%=rs("mh")%>">
				<input type="hidden" name="dwmc1" value="<%=rs("dwmc")%>">
				<input type="hidden" name="lsno" value=1>
				<input type="hidden" name="indbinf" value="add">
				<input type="submit" value=" 确定 ">
			</td>
		</tr>
		</form>
	</table>
<%
End Function
%>
<%
Function DocBakAdds(rs)
%>

		<form name="frm_docbak" id="frm_docbak" action="docbak_indb.asp" method="post" onsubmit='return docbak_checkinf();'>

		<table width="68%" cellpadding=2 cellspacing=0 class=xtable>
		<tr>
		<td class=rtd>订单号</td>
			<td class=ltd><%=rs("ddh")%></td>
		</tr>
			<tr>
		<td class=rtd>单位名称</td>
			<td class=ltd><%=rs("dwmc")%></td>
		</tr>
		<tr>
			<td class=rtd>所存盘号</td>
			<td class=ltd><input type="text" name="diskid" size="15"></td>
		</tr>
		<tr>
			<td class=rtd>备注</td>
			<td class=ltd><textarea name="bz" cols="60" rows="10"></textarea></td>
		</tr>

	<%
    Dim s_lsh,i,n
	i=1
	s_lsh=Trim(Request("s_lsh"))
	if s_lsh=rs("lsh") Then
		strsql="select *  from mtask where ddh = (select ddh from mtask where lsh='"&s_lsh&"') and  not(cp)"
	Else
		strsql="select *  from mtask where ddh='"&s_lsh&"' and  not(cp)"
	End If
	Rs.Close
	Rs.open strsql,Conn ,1,3
	if not Rs.BOF then
	Rs.MoveLast

    n=Rs.RecordCount
	end if
	Response.Write("该订单号有"&n&"个流水号记录准备存盘!")
	Rs.MoveFirst
	do while not Rs.eof
'Call TbTopic("添加流水号" & rs("lsh") & " 模具图纸存档信息")
%>
		<tr>
			<td class=ctd colspan="2">
			<input type="hidden" name="lsh<%=i%>" value="<%=rs("lsh")%>">
			<input type="hidden" name="ddh<%=i%>" value="<%=rs("ddh")%>">
			<input type="hidden" name="mh<%=i%>" value="<%=rs("mh")%>">
			<input type="hidden" name="dwmc<%=i%>" value="<%=rs("dwmc")%>">




<%
i=i+1
	Rs.MoveNext
	Loop
	 %>

		<input type="hidden" name="indbinf" value="add">
		<input type="hidden" name="lsno" value="<%=n%>">
		<input type="submit" value=" 确定 ">
		</td>
		</tr>

	  </table>
	  </form>

<%
	End Function
%>
<%
Function waitSave()
	Dim strAll
	strAll=request("disall")
	If strAll="" Then strAll="yes"
%>

	<%Call TbTopic("等待存档的模具")%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable>
		<tr>
			<th class=th>id</th>
			<th class=th>订单号</th>
			<th class=th>流水号</th>
			<th class=th>单位名称</th>
			<th class=th>断面名称</th>
			<th class=th>完成时间</th>
			<th class=th>模号</th>
			<%if ChkAble(7) then response.write("<td class=th>操作</td>")%>
		</tr>
		<%
			Dim i
			Set Rs = xjweb.Exec("select * from [mtask] where not(isnull(sjjssj)) and not(cp) order by id",1)
			i = 1
			do while not rs.eof
		%>
				<tr>
					<td class=ctd><%=i%></td>
					<td class=ctd><%=rs("ddh")%></td>
					<td class=ctd><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
					<td class=ctd><%=rs("dwmc")%></td>
					<td class=ctd><%=rs("dmmc")%></td>
					<td class=ctd><%=rs("sjjssj")%></td>
					<td class=ctd><%=ucase(rs("mh") & "-" &  rs("lsh"))%>
					</td>

	                <% response.write("<td class=ctd><a href=""docbak_add.asp?s_lsh="&rs("lsh")&""" onclick=""getUserSelect() ;"">存盘</a></td>")%>
				</tr>
			<%
				if i >= 20 and strAll <> "yes" then exit do
				i=i+1
				rs.movenext
			loop
			rs.close
			set rs = nothing
		%>

	</table><br>
	<%if strAll <> "yes" then response.write("<a href='?disall=yes'>显示所有</a>")%>


<%end function%>
<%function SelectSave (rs,s_lsh)
dim temp
temp =Request.Cookies("useroperation")
if rs("ddh")=s_lsh Then temp="batch"
if temp="batch" then
Call DocBakAdds(rs)
else
Call DocBakAdd(rs)
end if
end function
 %>