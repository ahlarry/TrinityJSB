<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble("1,2,3")
CurPage="系统通知 → 发布通知"
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
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>" align="center">
		<Tr><Td class=ctd height=300>
			<%Call InformAdd()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function InformAdd()
	Dim action
	action=request("action")
	select case action
		case "addindb"
			call inform_addindb()
		case "add"
			call inform_add()
		case else
			call inform_add()
	end select
End Function

Function inform_add()
	Call TbTopic("发布系统通知")
%>
	<table cellspacing=0 cellpadding=2 class=xtable align="center">
		<form id="addinform" action="?action=addindb" method="post" onsubmit="return chkinf();">
		<tr>
			<td class=rtd>通知主题</td>
			<td class=ltd><input type="text" name="informtopic" size="70"></td>
		</tr>
		<tr>
			<td class=rtd>通知内容</td>
			<td class=ltd><textarea name="informcontent" rows="15" cols="70"></textarea></td>
		</tr>
		<tr>
			<td class=ctd colspan="2"><input type="submit" value=" 提 交 "></td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		function chkinf()
		{
			if(document.all.informtopic.value==""){alert('请填写通知主题!');document.all.informtopic.focus();return false;}
			if(document.all.informcontent.value==""){alert('请填写通知内容!');document.all.informcontent.focus();return false;}
		}
	</script>
<%
end function

function inform_addindb()
	dim informtopic, informcontent, informlzr, informdate
	informcontent=""
	informtopic=trim(request("informtopic"))
	if trim(request("informcontent"))<>"" then informcontent=request("informcontent")
	informlzr=session("userName")
	informdate=now()
	if informtopic="" or informcontent="" then Call JsAlert("通知的主题和内容均不能为空!","") : exit function
	strSql="select * from ims_inform"
	call xjweb.Exec("", -1)
	set rs=server.createobject("adodb.recordset")
	rs.open strSql, conn, 1, 3
	rs.addnew
		rs("inform_topic")=informtopic
		rs("inform_content")=informcontent
		rs("inform_date")=informdate
		rs("inform_lzr")=informlzr
	rs.update
	Call JsAlert("通知发布成功! ","inform_dis.asp")
end function
%>