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

'������ʼ
Function chgpwd_dis()
	Call TbTopic("���Ĺ���Ա <font style=""color:#ff0000;"">"&session("admin")&"</font> ����")
%>
	<table cellspacing=0 cellpadding=3 class=xtable>
		<form action="?action=chgpwd_indb" method="post" onsubmit="return chkinf();">
		<tr>
			<td class=rtd>������</td>
			<td class=ltd><input type="password" name="oldpwd"></td>
		</tr>
		<tr>
			<td class=rtd>������</td>
			<td class=ltd><input type="password" name="newpwd"></td>
		</tr>
		<tr>
			<td class=rtd>��֤����</td>
			<td class=ltd><input type="password" name="verifypwd"></td>
		</tr>
		<tr>
			<td class=ctd colspan=2><input type="submit" value=" �� �� "></td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		function chkinf()
		{
			if(document.all.oldpwd.value==""){alert("������Ϊ��!");document.all.oldpwd.focus();return false;}
			if(document.all.newpwd.value==""){alert("������Ϊ��!");document.all.newpwd.focus();return false;}
			if(document.all.verifypwd.value==""){alert("��֤����Ϊ��!");document.all.verifypwd.focus();return false;}
		}
	</script>
<%
end function

function chgpwd_indb()
	dim oldpwd, newpwd, verifypwd
	oldpwd=trim(request("oldpwd"))
	newpwd=trim(request("newpwd"))
	verifypwd=trim(request("verifypwd"))
	if oldpwd="" or newpwd="" or verifypwd="" then call jsalert("�����������Ϊ��!","?action=chgpwd") : exit function
	strSql="select * from [ims_admin] where admin_name='"&session("admin")&"'"
	set rs=xjweb.exec(strSql, 1)
	if not(rs.eof or rs.bof) then
		if md5(oldpwd,16)<>rs("admin_pwd") then rs.close : call jsalert("�����벻��ȷ!","?action=chgpwd") : exit function
	else
		call jsalert("����Ȩ�ޱ�����!","index.asp") : exit function
	end if
	rs.close
	if newpwd<>verifypwd then call jsalert("��֤���벻��ȷ!","?action=chgpwd") : exit function
	strSql="update [ims_admin] set admin_pwd='"&md5(newpwd,16)&"' where admin_name='"&session("admin")&"'"
	call xjweb.exec(strSql, 0)
	call jsalert("����Ա������ĳɹ�!","admin_index.asp")
end function

function chgadmin_dis()
	Call TbTopic("���Ĺ���Ա <font style=""color:#ff0000;"">"&session("admin")&"</font> ����")
%>
	<table cellspacing=0 cellpadding=3 class=xtable>
		<form action="?action=chgadmin_indb" method="post" onsubmit="return chkinf();">
		<tr>
			<td class=rtd>����Ա����</td>
			<td class=ltd><input type="text" name="adminname" value="<%=session("admin")%>"></td>
		</tr>
		<tr>
			<td class=rtd>����Ա����</td>
			<td class=ltd><input type="password" name="adminpwd"></td>
		</tr>
		<tr>
			<td class=ctd colspan="2"><input type="submit" value=" �� �� "></td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		function chkinf()
		{
			if(document.all.adminname.value==""){alert("����Ա����Ϊ��!");document.all.adminname.focus();return false;}
			if(document.all.adminpwd.value==""){alert("���������Ա����!");document.all.adminpwd.focus();return false;}
		}
	</script>
<%
end function

function chgadmin_indb()
	dim adminname, adminpwd
	adminname=trim(request("adminname"))
	adminpwd=md5(trim(request("adminpwd")),16)
	if adminname="" or adminpwd="" then call jsalert("����Ա���ƺ����벻��Ϊ��!","?action=chgadmin") : exit function
	strSql="select * from [ims_admin] where admin_name='"&session("admin")&"'"
	set rs=xjweb.exec(strSql, 1)
	if not(rs.eof or rs.bof) then
		if rs("admin_pwd")<>adminpwd then rs.close : call jsalert("����Ա���벻��ȷ!","?action=chgadmin") : exit function
		if rs("admin_name")=adminname then rs.close : call jsalert("����Ա����û�б�!","?action=chgadmin") : exit function
	end if
	rs.close
	strSql="update [ims_admin] set admin_name='"&adminname&"' where admin_name='"&session("admin")&"'"
	call xjweb.exec(strSql, 0)
	session("admin")=adminname
	call jsalert("����Ա���Ƹ��ĳɹ�!","admin_index.asp")
end function
%>