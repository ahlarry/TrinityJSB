<!--#include file="include/conn.asp"-->
<!--#include file="include/md5.asp"-->
<%
Call ChkPageAble(1)
xjweb.header()
	dim action
	action=request("action")
	select case action
		case "chgadmin"
			call chgadmin_dis()
		case "chgadmin_indb"
			call chgadmin_indb()
		case "chgpwd"
			call chgpwd_dis()
		case "chgpwd_indb"
			call chgpwd_indb()
		case else
			response.write("action = " & action)
	end select
xjweb.footer()

'函数开始
Function chgpwd_dis()
	Call TbTopic("更改管理员 <font style=""color:#ff0000;"">"&session("admin")&"</font> 密码")
%>
	<table cellspacing=0 cellpadding=3 class=xtable>
		<form action="?action=chgpwd_indb" method="post" onsubmit="return chkinf();">
		<tr>
			<td class=rtd>旧密码</td>
			<td class=ltd><input type="password" name="oldpwd"></td>
		</tr>
		<tr>
			<td class=rtd>新密码</td>
			<td class=ltd><input type="password" name="newpwd"></td>
		</tr>
		<tr>
			<td class=rtd>验证密码</td>
			<td class=ltd><input type="password" name="verifypwd"></td>
		</tr>
		<tr>
			<td class=ctd colspan=2><input type="submit" value=" 更 改 "></td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		function chkinf()
		{
			if(document.all.oldpwd.value==""){alert("旧密码为空!");document.all.oldpwd.focus();return false;}
			if(document.all.newpwd.value==""){alert("新密码为空!");document.all.newpwd.focus();return false;}
			if(document.all.verifypwd.value==""){alert("验证密码为空!");document.all.verifypwd.focus();return false;}
		}
	</script>
<%
end function

function chgpwd_indb()
	dim oldpwd, newpwd, verifypwd
	oldpwd=trim(request("oldpwd"))
	newpwd=trim(request("newpwd"))
	verifypwd=trim(request("verifypwd"))
	if oldpwd="" or newpwd="" or verifypwd="" then call jsalert("密码项均不能为空!","?action=chgpwd") : exit function
	strSql="select * from [ims_admin] where admin_name='"&session("admin")&"'"
	set rs=xjweb.exec(strSql, 1)
	if not(rs.eof or rs.bof) then
		if md5(oldpwd,16)<>rs("admin_pwd") then rs.close : call jsalert("旧密码不正确!","?action=chgpwd") : exit function
	else
		call jsalert("您的权限被剥夺!","index.asp") : exit function
	end if
	rs.close
	if newpwd<>verifypwd then call jsalert("验证密码不正确!","?action=chgpwd") : exit function
	strSql="update [ims_admin] set admin_pwd='"&md5(newpwd,16)&"' where admin_name='"&session("admin")&"'"
	call xjweb.exec(strSql, 0)
	call jsalert("管理员密码更改成功!","admin_index.asp")
end function

function chgadmin_dis()
	Call TbTopic("更改管理员 <font style=""color:#ff0000;"">"&session("admin")&"</font> 名称")
%>
	<table cellspacing=0 cellpadding=3 class=xtable>
		<form action="?action=chgadmin_indb" method="post" onsubmit="return chkinf();">
		<tr>
			<td class=rtd>管理员名称</td>
			<td class=ltd><input type="text" name="adminname" value="<%=session("admin")%>"></td>
		</tr>
		<tr>
			<td class=rtd>管理员密码</td>
			<td class=ltd><input type="password" name="adminpwd"></td>
		</tr>
		<tr>
			<td class=ctd colspan="2"><input type="submit" value=" 更 改 "></td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		function chkinf()
		{
			if(document.all.adminname.value==""){alert("管理员名称为空!");document.all.adminname.focus();return false;}
			if(document.all.adminpwd.value==""){alert("请输入管理员密码!");document.all.adminpwd.focus();return false;}
		}
	</script>
<%
end function

function chgadmin_indb()
	dim adminname, adminpwd
	adminname=trim(request("adminname"))
	adminpwd=md5(trim(request("adminpwd")),16)
	if adminname="" or adminpwd="" then call jsalert("管理员名称和密码不能为空!","?action=chgadmin") : exit function
	strSql="select * from [ims_admin] where admin_name='"&session("admin")&"'"
	set rs=xjweb.exec(strSql, 1)
	if not(rs.eof or rs.bof) then
		if rs("admin_pwd")<>adminpwd then rs.close : call jsalert("管理员密码不正确!","?action=chgadmin") : exit function
		if rs("admin_name")=adminname then rs.close : call jsalert("管理员名称没有变!","?action=chgadmin") : exit function
	end if
	rs.close
	strSql="update [ims_admin] set admin_name='"&adminname&"' where admin_name='"&session("admin")&"'"
	call xjweb.exec(strSql, 0)
	session("admin")=adminname
	call jsalert("管理员名称更改成功!","admin_index.asp")
end function
%>