<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(1)
CurPage="IP管理 → IP管理"
strPage=""
Call FileInc(0,"js/ipman.js")
xjweb.header()
Call TopTable()

Dim action, strIP, strMac, strUser, strRemark
strIP=Trim(Request("ip"))
strMac=Trim(Request("Mac"))
strRemark=Request("remark")
strUser=Request("ipuser")

action=LCase(Request("action"))
Select Case action
	Case "ipsch"
		Call IpSearch()
	Case "macsch"
		Call MacSearch()
	Case "ipadd"
		Call IpAdd()
	Case "ipchange"
		Call IpChange()
	Case Else
		Call Main()
End Select

Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%Call techAdd()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

function techAdd()
%>
	<%Call TbTopic("内部IP管理系统")%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="80%" align="center">
		<form id=frm_ipman1 name=frm_ipman1 action=<%=Request.Servervariables("SCRIPT_NAME")%> method=post onSubmit='return chkipman1(this);'>
		<tr>
			<td class=th width=120>IP → MAC 查询</td>
			<td class=ltd>
				请输入IP:<input type=text name=ip>
				<input type="submit" value=" 查询 ">
			</td>
		</tr>
		<input type="hidden" name="action" value="ipsch">
		</form>
	</Table>
	<%Response.Write(XjLine(10,"100%",""))%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="80%" align="center">
		<form id=frm_ipman2 name=frm_ipman2 action=<%=Request.Servervariables("SCRIPT_NAME")%> method=post onSubmit='return chkipman2(this);'>
		<tr>
			<td class=th width=120>MAC → IP 查询</td>
			<td class=ltd>
				请输入MAC:<input type=text name=mac>
				<input type="submit" value=" 查询 ">
			</td>
		</tr>
		<input type="hidden" name="action" value="macsch">
		</form>
	</Table>
	<%Response.Write(XjLine(10,"100%",""))%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="80%" align="center">
		<form id=frm_ipman3 name=frm_ipman3 action=<%=Request.Servervariables("SCRIPT_NAME")%> method=post onSubmit='return chkipman3(this);'>
		<tr>
			<td class=th width=120>Add IP</td>
			<td class=ltd>
				添加IP:<input type=text name=ip>
				<input type="submit" value=" 添加 ">
			</td>
		</tr>
		<input type="hidden" name="action" value="ipadd">
		</form>
	</Table>
	<%Response.Write(XjLine(10,"100%",""))%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="80%" align="center">
		<form id=frm_ipman4 name=frm_ipman4 action=<%=Request.Servervariables("SCRIPT_NAME")%> method=post onSubmit='return chkipman4(this);'>
		<tr>
			<td class=th width=120>Change IP</td>
			<td class=ltd>
				<Table class=xtable cellspacing=0 cellpadding=3 width="100%">
					<Tr>
						<Td class=rtd>请输入IP:</Td>
						<Td class=ltd><input type=text name="ip" value=<%=strIP%>> <input type="submit" name="submit" value=" 查询 "></Td>
					</Tr>
					<Tr>
						<Td class=rtd>MAC:</Td>
						<Td class=ltd><input type=text name="MAC" value="<%=strMac%>"></Td>
					</Tr>
					<Tr>
						<Td class=rtd>使用者:</Td>
						<Td class=ltd><input type=text name="ipuser" value="<%=strUser%>"></Td>
					</Tr>
					<Tr>
						<Td class=rtd>Remarks:</Td>
						<Td class=ltd><textarea name=remark cols=50 rows=5><%=strRemark%></textarea></Td>
					</Tr>
					<Tr>
						<Td class=rtd>&nbsp;</Td>
						<Td class=ctd><input type="submit" name="submit" value=" 更改 "></Td>
					</Tr>
				</Table>
			</td>
		</tr>
		<input type="hidden" name="action" value="ipchange">
		</form>
	</Table>
	<%Response.Write(XjLine(10,"100%",""))%>
<%
end function

Function IpSearch()
	Dim iip, tip
	tip=Split(strip,".")
	If ubound(tip)<>3 Then Call JsAlert("IP地址输入错误!" , "") : Exit Function
	for i=0 to ubound(tip)
		If Not IsNumeric(tip(i)) Then Call JsAlert("IP地址只能为数字!","") : Exit Function
		'strip(i)=CInt(strip(i))
		If tip(i)<0 Or tip(i)>255 Then Call JsAlert("IP地址数值应在0-255之间!\n请重新输入!","") : Exit Function
	next
	'iip=tip(0)*256*256*256 +tip(1)*256*256 + tip(2)*256 + tip(3)
	strSql="select * from ims_ip where ip='" & strip & "'"
	set Rs=xjweb.Exec(strSql, 1)
	If Rs.Eof Then Call JsAlert("IP地址:"&strip&" 暂时无人使用!","") : Exit Function
	Call JsAlert("IP地址:" & strip &"\n使用者:"&Rs("uName")&"\nMAC:"&Rs("mac")&"","")
End Function

Function IpAdd()
	Dim iip, tip
	tip=Split(strip,".")
	If ubound(tip)<>3 Then Call JsAlert("IP地址输入错误!" , "") : Exit Function
	for i=0 to ubound(tip)
		If Not IsNumeric(tip(i)) Then Call JsAlert("IP地址只能为数字!","") : Exit Function
		'strip(i)=CInt(strip(i))
		If tip(i)<0 Or tip(i)>255 Then Call JsAlert("IP地址数值应在0-255之间!\n请重新输入!","") : Exit Function
	next
	'iip=tip(0)*256*256*256 +tip(1)*256*256 + tip(2)*256 + tip(3)
	strSql="select * from ims_ip where ip='" & strip & "'"
	set Rs=xjweb.Exec(strSql, 1)
	If Not Rs.Eof Then Call JsAlert("IP地址:"&strip&" 已存在!","") : Exit Function
	Rs.Close
	strSql="insert into ims_ip (ip) values ('" & strip & "')"
	Call xjweb.Exec(strSql, 0)
	Call JsAlert("IP地址:" & strip &"\n添加成功！","")
End Function

Function IpChange()
	Dim iip, tip
	tip=Split(strip,".")
	If ubound(tip)<>3 Then Call JsAlert("IP地址输入错误!" , "") : Exit Function
	for i=0 to ubound(tip)
		If Not IsNumeric(tip(i)) Then Call JsAlert("IP地址只能为数字!","") : Exit Function
		'strip(i)=CInt(strip(i))
		If tip(i)<0 Or tip(i)>255 Then Call JsAlert("IP地址数值应在0-255之间!\n请重新输入!","") : Exit Function
	next
	strSql="select * from ims_ip where ip='" & strip & "'"
	set Rs=xjweb.Exec(strSql, 1)
	If Rs.Eof Then Call JsAlert("请先添加:\nIP地址:"&strip&" !","") : Exit Function

	If Trim(Request("submit"))="查询" Then
		strMac=Rs("mac")
		strRemark=Rs("remark")
		strUser=Rs("uName")
		Rs.Close
		Call Main()
	Else
		If strMac="" Then Call JsAlert("请输入ＭＡＣ地址!","") : Exit Function
		If strRemark="" Then strRemark=" "
		If strUser="" Then strUser=" "
		strSql="update [ims_ip] set mac='"&strMac&"', remark='"&strRemark&"' , uName='"&strUser&"' where ip='" & strip & "'"
		Call xjweb.Exec(strSql,0)
		Call JsAlert("ＩＰ及ＭＡＣ地址更改成功!","") : Exit Function
	End If
End Function

Function MacSearch()
	If strMac="" Then Call JsAlert("请输入ＭＡＣ!" , "") : Exit Function
	strSql="select * from [ims_ip] where mac='"&strMac&"'"
	set Rs=xjweb.Exec(strsql, 1)
	If Rs.Eof Then Call JsAlert("此ＭＡＣ地址暂未登入，请添加!" , "") : Exit Function

	Dim tip
	tip=Rs("ip")
	strUser=Rs("uName")
	Call JsAlert("Mac地址: " & strMac & " \n IP地址:" & tip &"\n使用者:"&strUser&"","")
End Function
%>