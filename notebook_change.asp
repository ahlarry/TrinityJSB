<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="系统留言 →  更改留言"
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
			<%Call NoteBookChange()%>
		</Td></Tr>
	</Table>
<%
End Sub

Function NoteBookChange()
	Dim iid
	iid = clng(request("id"))
	If not isnumeric(iid) Then
		Call JsAlert("请从正确的入口进入,拜托!","notebook.asp")
	Else
		strSql ="select * from [notebook] where id = "&iid&""
		Set Rs=xjweb.Exec(strSql, 1)
		If Rs.Eof Or Rs.Bof Then 
			Call JsAlert("留言不存在，可能已经删除！","notebook.asp")
		ElseIf Session("userName") <> rs("username") then 
			Call JsAlert("您无权更改 "&rs("username")&" 的留言","notebook.asp")
		Else
			Call NoteBook_Change(rs)
		End If
		Rs.Close
	End If
End Function
%>

<%
Function NoteBook_Change(rs)
	Call TbTopic("更改留言")
%>
	<table width="80%" cellpadding=3 cellspacing="0" class=xtable>
		<form action="notebook_indb.asp" method="post" onsubmit="return checkinf();">
		<tr><td class=ltd>留言内容:</Td></Tr>
		<tr><td class=ctd><textarea name="lylr" cols="90" rows="15"><%=rs("content")%></textarea></td></tr>
			<input type="hidden" name="id" value="<%=rs("id")%>">
			<input type="hidden" name="indbinf" value="change">
		<tr><td class=ctd><input type="submit" value=" 提交更改 "></td></tr>
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
		if (trim(document.all.lylr.value)==""){alert("留言内容不能为空！\n");document.all.lylr.focus();document.all.lylr.value="";return false;}
	}
</script>