<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble("1,2,3")
CurPage="ϵͳ֪ͨ �� ����֪ͨ"
strPage="inform"
'Call FileInc(0, "js/docbak.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%Call InformChange()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function InformChange()
	Dim action
	action=request("action")
	select case action
		case "changeindb"
			call inform_changeindb()
		case "change"
			call inform_change()
		case "delete"
			call inform_delete()
		case Else
			Call JsAlert("�����ȷ����ڽ��룡","inform_dis.asp")
	end select
End Function

Function inform_change()
	Dim iid
	iid=Request("id")
	if not(isnumeric(iid)) then iid=0
	iid=clng(iid)
	strSql="select * from [ims_inform] where id="&iid&""
	set rs=xjweb.Exec(strSql, 1)
	if not(rs.eof or rs.bof) then
		Call TbTopic("����֪ͨ")
%>
		<table cellspacing=0 cellpadding=2 class=xtable>
			<form id="addinform" action="?action=changeindb" method="post" onsubmit="return chkinf();">
			<tr>
				<td class=rtd>֪ͨ����</td>
				<td class=ltd><input type="text" name="informtopic" size="70" value=<%=rs("inform_topic")%>></td>
			</tr>
			<tr>
				<td class=rtd>֪ͨ����</td>
				<td class=ltd><textarea name="informcontent" rows="15" cols="70"><%=rs("inform_content")%></textarea></td>
			</tr>
			<tr>
				<td class=ctd colspan="2"><input type="submit" value=" �� �� "></td>
			</tr>
			<input type="hidden" name="id" value=<%=rs("id")%>>
			</form>
		</table>
		<script language="javascript">
			function chkinf()
			{
				if(document.all.informtopic.value==""){alert('����д֪ͨ����!');document.all.informtopic.focus();return false;}
				if(document.all.informcontent.value==""){alert('����д֪ͨ����!');document.all.informcontent.focus();return false;}
			}
		</script>
<%
	else
		Call JsAlert("�����ȷ��ڽ���!","")
	end if
End Function

Function inform_changeindb()
	dim informtopic, informcontent, informlzr, informdate, informid
	informcontent=""
	informtopic=trim(request("informtopic"))
	if trim(request("informcontent"))<>"" then informcontent=request("informcontent")
	informlzr=session("userName")
	informdate=now()
	informid=request("id")
	if not(isnumeric(informid)) then informid=0
	if informtopic="" or informcontent="" or informid=0 then Call JsAlert("֪ͨ����������ݾ�����Ϊ��,��ֻ�ܴ���ȷ����ڽ���!","") : Exit Function
	strSql="select * from [ims_inform] where id="&informid&""
	call xjweb.Exec("", -1)
	rs.open strSql, conn, 1, 3
		rs("inform_topic")=informtopic
		rs("inform_content")=informcontent
		rs("inform_date")=informdate
		rs("inform_lzr")=informlzr
	rs.update
	Call JsAlert("֪ͨ���ĳɹ�!","inform_dis.asp?id="&informid&"")
End Function

Function inform_delete()
	Dim iid
	iid=request("id")
	if not(isnumeric(iid)) then iid=0
	if iid=0 then Call JsAlert("�����ȷ����ڽ���!","") : Exit Function
	strSql="delete from [ims_inform] where id="&iid&""
	call xjweb.Exec(strSql, 0)
	Call JsAlert("֪ͨɾ���ɹ�!","inform_dis.asp")
End Function
%>