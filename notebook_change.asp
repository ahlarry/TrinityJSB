<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="ϵͳ���� ��  ��������"
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
		Call JsAlert("�����ȷ����ڽ���,����!","notebook.asp")
	Else
		strSql ="select * from [notebook] where id = "&iid&""
		Set Rs=xjweb.Exec(strSql, 1)
		If Rs.Eof Or Rs.Bof Then 
			Call JsAlert("���Բ����ڣ������Ѿ�ɾ����","notebook.asp")
		ElseIf Session("userName") <> rs("username") then 
			Call JsAlert("����Ȩ���� "&rs("username")&" ������","notebook.asp")
		Else
			Call NoteBook_Change(rs)
		End If
		Rs.Close
	End If
End Function
%>

<%
Function NoteBook_Change(rs)
	Call TbTopic("��������")
%>
	<table width="80%" cellpadding=3 cellspacing="0" class=xtable>
		<form action="notebook_indb.asp" method="post" onsubmit="return checkinf();">
		<tr><td class=ltd>��������:</Td></Tr>
		<tr><td class=ctd><textarea name="lylr" cols="90" rows="15"><%=rs("content")%></textarea></td></tr>
			<input type="hidden" name="id" value="<%=rs("id")%>">
			<input type="hidden" name="indbinf" value="change">
		<tr><td class=ctd><input type="submit" value=" �ύ���� "></td></tr>
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
		if (trim(document.all.lylr.value)==""){alert("�������ݲ���Ϊ�գ�\n");document.all.lylr.focus();document.all.lylr.value="";return false;}
	}
</script>