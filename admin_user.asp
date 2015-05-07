<!--#include file="include/conn.asp"-->
<!--#include file="include/md5.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/depart_dbinf.asp"-->
<%
Call ChkAdminAble()
xjweb.Header()
	Dim action
	action=Request("action")
	Select Case action
		Case "add"
			Call AddUser()
		Case "delete"
			Call DeleteUser()
		Case "change"
			Call ChangeUser()
		Case Else
			Call Main()
			If action<>"" Then Response.Write("<br><br><br>" & action)
	End Select
xjweb.Footer()

'函数开始
Sub Main()
	Call TbTopic(web_info(0) & "--用户信息设置")
	Response.Write(xjLine(1,"100%",web_info(12)))
%>
	<Table cellspacing=0 cellpadding=4 border=0 width=500>
		<Tr><Td>
			<%Call DisAddUser()%>
		</Td></Tr>
		<Tr><Td>
			<%Call DisDeleteUser()%>
		</Td></Tr>
		<Tr><Td>
			<%Call DisChangeUser()%>
		</Td></Tr>
	</Table>
<%
End Sub

Sub DisAddUser()
%>
	<Table cellspacing=0 cellpadding=4 class=xtable width="100%">
		<Form name="frm_adduser" action="<%=request.servervariables("script_name")%>" method=post onsubmit="return userAddChk();">
		<Tr>
			<Td class=th width=100>添加用户</Td>
			<Td class=ltd width=*><input type="text" name="userName"></Td>
			<Td class=ctd width=100><input type="submit" value=" 添 加 "></Td>
		</Tr>
		<input type="hidden" name="action" value="add">
		</Form>
	</Table>
	<script language="javascript">
		function userAddChk()
		{
			 if (document.frm_adduser.userName.value=="")
				{ alert("请输入待添加的 用户名称 ！"); document.frm_adduser.userName.focus(); return false;}
		}
	</script>
<%
End Sub

Sub DisDeleteUser()
%>
	<Table cellspacing=0 cellpadding=4 class=xtable width="100%">
		<Form name="frm_deleteuser" action="<%=request.servervariables("script_name")%>" method=post onsubmit="return userChangeChk();">
		<Tr>
			<Td class=th width=100>删除用户</Td>
			<Td class=ltd width=*>
				<Select name="userName">
					<option></option>
					<%for i=0 to ubound(c_alluser)%>
						<option value="<%=c_alluser(i)%>"><%=c_alluser(i)%></option>
					<%next%>
				</Select>
			</Td>
			<Td class=ctd width=100><input type="submit" value=" 删 除 "></Td>
		</Tr>
		<input type="hidden" name="action" value="delete">
		</Form>
	</Table>
	<script language="javascript">
		function userChangeChk()
		{
			if (document.frm_deleteuser.userName.value=="")
			{ alert("请选择待删除的 用户 ！"); document.frm_deleteuser.userName.focus(); return false;}
			else
			{return confirm("用户删除后将不能恢复！\n确信删除用户 【" + document.frm_deleteuser.userName.value + "】 吗！");}
		}
	</script>
<%
End Sub

Sub DisChangeUser()
%>
	<Table cellspacing=0 cellpadding=4 class=xtable width="100%">
		<Tr>
			<Td class=th width=100>更改用户信息</Td>
			<Td class=ltd width=*>
				<Table cellspacing=0 cellpadding=2 class=xtable width="100%">
					<Form name="frm_selectuser" action="<%=request.servervariables("script_name")%>" method=post>
					<Tr><Td class=ltd>
						<Select name="userName" onchange="location.href='?userName='+this.value;">
							<option></option>
							<%for i=0 to ubound(c_alluser)%>
								<%If Request("userName")=c_alluser(i) Then%>
									<option value="<%=c_alluser(i)%>" selected><%=c_alluser(i)%></option>
								<%Else%>
									<option value="<%=c_alluser(i)%>"><%=c_alluser(i)%></option>
								<%End If%>
							<%next%>
						</Select>&nbsp;&nbsp;如没有自动跳转请点击
						<input type="submit" value=" 选 择 ">
					</Td></Tr>
					</Form>
				</Table>
				<br>

				<%
				If Trim(Request("userName"))<>"" Then
					Dim strUser
					strUser=Trim(Request("userName"))
					strSql="select * from [ims_user] where user_name='"&strUser&"'"
					Set Rs=xjweb.Exec(strSql, 1)
					If Rs.Eof Or Rs.Bof Then
						Call JsAlert("用户不存在啊！","admin_user.asp")
					Else
				%>
						<Table cellspacing=0 cellpadding=2 class=xtable width="100%">
						<Form name="frm_changeuser" action="<%=request.servervariables("script_name")%>" method=post>
							<Tr><Td class=th colspan=2>更改 <%=strUser%> 信息</Td></Tr>
							<Tr>
							<Td class=rtd width=60>重置密码</Td>
							<Td class=ltd><input type=checkbox name="czmm" value=true>是</Td>
							<Tr>
								<Td class=rtd width=60>权限</Td>
								<Td class=ltd>
									<input type=checkbox name="able1" value=true <%if chkuser(1) then%> checked <%end if%>>管理员<br>
									<input type=checkbox name="able2" value=true <%if chkuser(2) then%> checked <%end if%>>厂长<br>
									<input type=checkbox name="able3" value=true <%if chkuser(3) then%> checked <%end if%>>经理<br>
									<input type=checkbox name="able4" value=true <%if chkuser(4) then%> checked <%end if%>>组长<br>
									<input type=checkbox name="able5" value=true <%if chkuser(5) then%> checked <%end if%>>组员<br>
									<input type=checkbox name="able6" value=true <%if chkuser(6) then%> checked <%end if%>>模具调试员<br>
									<input type=checkbox name="able7" value=true <%if chkuser(7) then%> checked <%end if%>>图档管理员<br>
									<input type=checkbox name="able8" value=true <%if chkuser(8) then%> checked <%end if%>>服务技术员<br>
									<input type=checkbox name="able9" value=true <%if chkuser(9) then%> checked <%end if%>>编程技术员<br>
									<input type=checkbox name="able10" value=true <%if chkuser(10) then%> checked <%end if%>>副主任<br>
									<input type=checkbox name="able11" value=true <%if chkuser(11) then%> checked <%end if%>>品管部<br>
									<input type=checkbox name="able12" value=true <%if chkuser(12) then%> checked <%end if%>>管理部<br>
									<input type=checkbox name="able13" value=true <%if chkuser(13) then%> checked <%end if%>>网络管理员<br>
								</td>
							</Tr>
							<tr>
								<td class=rtd>组</td>
								<td class=ltd>
									<Select name="user_group">
										<%for i=0 to 8%>
											<option value=<%=i%> <%If i=Rs("user_group") then response.write(" selected")%>><%=i%></option>
										<%next%>
									</Select>&nbsp;&nbsp;第5组负责调试任务的分配
								</td>
							</tr>
							<tr>
								<td class=rtd>部门</td>
								<td class=ltd>
									<select name="userDepart">
										<%if isnull(rs("user_depart")) then%><option></option><%end if%>
										<%for i=0 to ubound(c_depart)%>
											<option value='<%=c_depart(i)%>' <%if rs("user_depart")=c_depart(i) then %> selected<%end if%>><%=c_depart(i)%></option>
										<%next%>
									</select>
								</td>
							</tr>
							<tr>
								<td class=rtd>用户拼音</td>
								<td class=ltd><input type=text name="user_Spelling" value=<%=rs("user_Spelling")%>></td>
							</tr>
							<tr>
								<td class=rtd>拼音缩写</td>
								<td class=ltd><input type=text name="user_abb" value=<%=rs("user_abb")%>></td>
							</tr>
							<tr>
								<td class=rtd>IP</td>
								<td class=ltd><input type=text name="user_ip" value=<%=rs("user_ip")%>></td>
							</tr>
							<Tr><Td colspan=2 class=ctd>
								<input type="hidden" name="userName" value="<%=strUser%>">
								<input type="hidden" name="action" value="change">
								<input type="submit" value=" 更 改 ">
							</Td></Tr>
						</Form>
						</Table>
						<%
					End IF
				End IF%>
			</Td>
		</Tr>
	</Table>
<%
End Sub

'添加用户
Function AddUser()
	Dim strName
	strName=Trim(Request("userName"))
	If strName="" Then Call JsAlert("请输入要添加的用户名!","")
	strSql="Select * from [ims_user] where user_name='"&strName&"'"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.Eof Or Rs.Bof Then
		strSql="Insert into [ims_user] ([user_name]) values ('"&strName&"') "
		Call xjweb.Exec(strSql,0)
		Call JsAlert("用户 【" & strName & "】 添加成功！","admin_user.asp?userName=" & strName)
	Else
		Call JsAlert("用户 【" & strName & "】 已存在！","")
	End If
	Rs.Close
End Function

'删除用户
Function DeleteUser()
	Dim strName
	strName=Trim(Request("userName"))
	If strName="" Then Call JsAlert("请输入要删除的用户!","")
	strSql="delete from [ims_user] where user_name='"&strName&"'"
	Call xjweb.Exec(strSql, 0)
	Call JsAlert("用户〖"&strName&" 〗删除成功!","admin_user.asp?action=")
End Function

'设置用户
Function ChangeUser()
	Dim struserName, struserAble, struserDepart, struserip, struserSpelling, struserabb, strczmm
	Dim struserGroup
	strczmm=""
	strusername=Trim(Request("userName"))
	struserSpelling=Trim(Request("user_Spelling"))
	struserip=Trim(Request("user_ip"))
	struserabb=trim(Request("user_abb"))
	struserGroup=Cint(Request("user_group"))

	If Request("czmm") Then strczmm="1"
	If Request("able1") Then struserAble="1" Else struserAble="0"
	If Request("able2") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able3") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able4") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able5") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able6") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able7") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able8") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able9") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able10") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able11") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able12") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able13") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able14") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"
	If Request("able15") Then struserAble=struserAble & "1" Else struserAble=struserAble & "0"

	struserdepart=Trim(Request("userDepart"))
	strSql="select * from [ims_user] where [user_name]='"&struserName&"'"
	Call xjweb.Exec(strSql,-1)
	'set rs=server.createobject("adodb.recordset")
	Rs.open strSql,conn,1,3
		rs("user_able")=struserAble
		if strczmm="1" Then rs("user_pwd")=md5("8888",16)
		if struserdepart<>"" Then rs("user_depart")=struserdepart
		if struserSpelling<>"" then rs("user_Spelling")=struserSpelling
		if struserabb<>"" then rs("user_abb")=struserabb
		If IsNumeric(strusergroup) Then Rs("user_group")=strusergroup
	rs.update
	Call JsAlert("用户 【"&struserName&"】 更改成功!","admin_user.asp?action=")
End Function
%>