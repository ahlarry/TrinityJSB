<!--#include file="include/conn.asp"-->
<!--#include file="include/md5.asp"-->
<%
Call ChkPageAble(0)
CurPage="�û����� �� �����û���Ϣ"
strPage="uctrl"
'Call FileInc(0, "js/mtask.js")
xjweb.header()
Call TopTable()

Dim action
action=""
action=Request("action")
Select Case action
	Case "change"
		Call change()
	Case "dis"
		Call Main()
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
			<%Call UctrlDis()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function UctrlDis()
	Call TbTopic("�����û�����")
%>
	<table cellpadding=4 cellspacing=0 class=xtable width="80%" align="center">
		<form name="myform" onsubmit="return chkinf();" action="?action=change" method="post">
		<Tr>
			<Td class=th>�û���</Td>
			<Td class=th>�û�����</Td>
			<Td class=th>�û�ͷ��</Td>
			<Td class=th>�û��ǳ�</Td>
		</Tr>
		<tr>
			<td class=ctd rowspan="2"><b><%=Session("userName")%></b></td>
			<td class=ctd>
				<table cellpadding=2 cellspacing=0 class=xtable width=200>
					<tr>
						<td class=rtd width="50">������</td>
						<td class=ltd><input type="password" name="oldpwd"></td>
					</tr>
					<tr>
						<td class=rtd width="50">������</td>
						<td class=ltd><input type="password" name="newpwd"></td>
					</tr>
					<tr>
						<td class=rtd width="50">��֤����</td>
						<td class=ltd><input type="password" name="verifypwd"></td>
					</tr>
				</table>
			</td>
			<td class=ctd><img id="preimg" src="<%=web_info(2)%>images/face/<%=Session("userFace")%>" width="80" height="80"><br><br>
				<input type="hidden" name="userhead"><input type="button" value="ѡ��ͷ��" onclick="showhead();">
			</td>
			<td class=ctd><input type="text" name="usernick" value="<%=Session("userNick")%>"></td>
		</tr>
		<tr>
			<td class=ctd><input type="submit" name="submit" value=" �������� "></td>
			<td class=ctd><input type="submit" name="submit" value=" ����ͷ�� "></td>
			<td class=ctd><input type="submit" name="submit" value=" �����ǳ� "></td>
		</tr>
		</form>
	</table>

	<div id="divhead" style="left:20;top:20;width=320;height:320;position:absolute;background-color:#eeeeee;border:1;z-index:1;display:none;">
		<table cellspacing=0 cellpadding=2 class=xtable align="center">
			<tr>
			<%for i=0 to 25%>
				<td class=ctd><img src="<%=web_info(2)%>images/face/<%=i%>.gif" width=40 height=40 onclick="document.all.userhead.value='<%=i%>.gif';document.all.preimg.src='<%=web_info(2)%>images/face/<%=i%>.gif';document.all.divhead.style.display='none';"></td>
				<%if (i+1) mod 8 = 0 then%></tr><tr><%end if%>
			<%next%>
			</tr>
		</table>
		���ѡ��
	</div>
	<script language="javascript">
		function chkinf()
		{
			if(document.all.newpwd.value!='')
			{
				if(document.all.oldpwd.value==''){alert("Ҫ��������������ԭ����!");document.all.oldpwd.focus();return false;}
				if(document.all.verifypwd.value==''){alert("��������֤����!");document.all.verifypwd.focus();return false;}
				if(document.all.newpwd.value!=document.all.verifypwd.value){alert("��֤���벻��ȷ");document.all.verifypwd.focus();return false;}
			}
		}
		function showhead()
		{
			document.all.divhead.style.left=(document.body.scrollWidth-280)/2;
			document.all.divhead.style.top=event.clientY-335;
			if(document.all.divhead.style.display=="none")
				document.all.divhead.style.display='';
			else
				document.all.divhead.style.display='none';
		}
	</script>
<%
end function

function change()
	Dim strChg
	strChg=Trim(Request("submit"))
	Select Case strChg
		Case "��������"
			Dim oldpwd, newpwd, verifypwd
			oldpwd=trim(request("oldpwd"))
			newpwd=Trim(request("newpwd"))
			verifypwd=trim(request("verifypwd"))
			if trim(newpwd)="" then Call JsAlert("��û����������������ǿհ׷�!","") : Exit Function
			if verifypwd<>trim(newpwd) then Call JsAlert("'��֤���벻��ȷ!","") : Exit Function
			strSql="select * from [ims_user] where user_name='"&Session("userName")&"'"
			set rs=xjweb.Exec(strSql, 1)
			if isnull(rs("user_pwd")) or rs("user_pwd")=md5(oldpwd,16) then
				strSql="update [ims_user] set user_pwd='"&md5(verifypwd,16)&"' where user_name='"&Session("userName")&"'"
				call xjweb.Exec(strSql, 0)
				Call JsAlert("������ĳɹ�!","uctrl_changeinf.asp?action=main") : Exit Function
			else
				Call JsAlert("ԭ���벻��ȷ!����������!","") : Exit Function
			end if
		Case "����ͷ��"
			dim strimg
			strimg=request("userhead")
			If strimg="" Then Call JsAlert("ͷ��û�и���!","") : Exit Function
			strSql="update [ims_user] set user_face='"&strimg&"' where user_name='"&session("userName")&"'"
			call xjweb.Exec(strSql, 0)
			Session("userFace")=strimg
			Call JsAlert("ͷ����ĳɹ�!","?action=main")
		Case "�����ǳ�"
			dim strusernick
			strusernick=trim(request("usernick"))
			strSql="select * from [ims_user] where user_nick='"&strusernick&"' and user_name<>'"&Session("userName")&"'"
			set rs=xjweb.Exec(strSql, 1)
			if rs.eof or rs.bof then
				strSql="update [ims_user] set user_nick='"&strusernick&"' where user_name='"&Session("userName")&"'"
				call xjweb.Exec(strSql, 0)
				session("userNick")=strusernick
				Call JsAlert("�û��ǳƸ��ĳɹ�!","?action=main")
			else
				Call JsAlert("�ǳ� "&strusernick&" �ѳɱ��˵��!","") : exit function
			end if
			rs.close
		Case Else
	End Select
	response.write "<script language='javascript'>location.href='?action=main';</script>"
End Function
%>
