<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="ϵͳ���� ��  �ظ�����"
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
		Call JsAlert("�����ȷ����ڽ���,����!","notebook.asp")
	else
		strSql ="select * from [notebook] where id = "&iid&""
		Set Rs=xjweb.Exec(strSql, 1)
		If ChkAble("1,2,3") then
			Call Notebook_Reply(rs)
		Else
			Call JsAlert("����Ȩ�ظ� "&rs("username")&" ������?","notebook.asp")
		End If
	End if
End Function
%>

<%
Function Notebook_Reply(rs)
	Call TbTopic("�ظ�����")
%>
	<table width="68%" cellpadding=3 cellspacing="0" class=xtable>
		<form action="notebook_indb.asp" method="post" onsubmit="return checkinf();">
		<tr><td class=th><%=rs("username")%>������:</td></tr>
		<tr><td class=ltd><%=xjweb.htmltocode(rs("content"))%></td></tr>
		<tr><td class=th>���Ļظ�:</td></tr>
		<tr><td class=ltd><textarea name="hf" cols="90" rows="10"></textarea></td></tr>
			<input type="hidden" name="id" value="<%=rs("id")%>">
			<input type="hidden" name="indbinf" value="reply">
		<tr><td class=ctd><input type="submit" value=" �ύ�ظ� "></td></tr>
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
			//����ִ���ߵ�һ���ַ�Ϊ�ո� 
			str = str.slice(1);//���ո���ִ���ȥ�� 
			//��һ��Ҳ�ɸĳ� str = str.substring(1, str.length); 
			str = lTrim(str); //�ݹ���� 
		} 
	return str; 
	} 

	//ȥ���ִ��ұߵĿո� 
	function rTrim(str) 
	{ 
		var iLength; 

		iLength = str.length; 
		if (str.charAt(iLength - 1) == " ") 
		{ 
			//����ִ��ұߵ�һ���ַ�Ϊ�ո� 
			str = str.slice(0, iLength - 1);//���ո���ִ���ȥ�� 
			//��һ��Ҳ�ɸĳ� str = str.substring(0, iLength - 1); 
			str = rTrim(str); //�ݹ���� 
		} 
		return str; 
	} 

	//ȥ���ִ����ߵĿո� 
	function trim(str) 
	{ 
		return lTrim(rTrim(str)); 
	} 


	function checkinf()
	{
		if (trim(document.all.hf.value)==""){alert("�ظ����ݲ���Ϊ�գ�\n");document.all.hf.focus();document.all.hf.value="";return false;}
	}
</script>