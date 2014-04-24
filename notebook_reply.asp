<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="系统留言 →  回复留言"
strPage="notebook"
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
			<%Call NoteBookReply()%>
		</Td></Tr>
	</Table>
<%
End Sub

Function NoteBookReply()
	Dim iid
	iid = clng(request("id"))
	if not isnumeric(iid) then
		Call JsAlert("请从正确的入口进入,拜托!","notebook.asp")
	else
		strSql ="select * from [notebook] where id = "&iid&""
		Set Rs=xjweb.Exec(strSql, 1)
		If ChkAble("1,2,3") then
			Call Notebook_Reply(rs)
		Else
			Call JsAlert("您无权回复 "&rs("username")&" 的留言?","notebook.asp")
		End If
	End if
End Function
%>

<%
Function Notebook_Reply(rs)
	Call TbTopic("回复留言")
%>
	<table width="68%" cellpadding=3 cellspacing="0" class=xtable>
		<form action="notebook_indb.asp" method="post" onsubmit="return checkinf();">
		<tr><td class=th><%=rs("username")%>的留言:</td></tr>
		<tr><td class=ltd><%=xjweb.htmltocode(rs("content"))%></td></tr>
		<tr><td class=th>您的回复:</td></tr>
		<tr><td class=ltd><textarea name="hf" cols="90" rows="10"></textarea></td></tr>
			<input type="hidden" name="id" value="<%=rs("id")%>">
			<input type="hidden" name="indbinf" value="reply">
		<tr><td class=ctd><input type="submit" value=" 提交回复 "></td></tr>
		</form>
	</table>
	<br>
<%
end function
%>
<script language="javascript">
	function lTrim(str) 
	{ 
		if (str.charAt(0) == " ") 
		{ 
			//如果字串左边第一个字符为空格 
			str = str.slice(1);//将空格从字串中去掉 
			//这一句也可改成 str = str.substring(1, str.length); 
			str = lTrim(str); //递归调用 
		} 
	return str; 
	} 

	//去掉字串右边的空格 
	function rTrim(str) 
	{ 
		var iLength; 

		iLength = str.length; 
		if (str.charAt(iLength - 1) == " ") 
		{ 
			//如果字串右边第一个字符为空格 
			str = str.slice(0, iLength - 1);//将空格从字串中去掉 
			//这一句也可改成 str = str.substring(0, iLength - 1); 
			str = rTrim(str); //递归调用 
		} 
		return str; 
	} 

	//去掉字串两边的空格 
	function trim(str) 
	{ 
		return lTrim(rTrim(str)); 
	} 


	function checkinf()
	{
		if (trim(document.all.hf.value)==""){alert("回复内容不能为空！\n");document.all.hf.focus();document.all.hf.value="";return false;}
	}
</script>