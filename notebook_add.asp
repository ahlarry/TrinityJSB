<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="ϵͳ���� ��  ǩд����"
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
			<%Call NoteBookAdd()%>
		</Td></Tr>
	</Table>
<%
End Sub

Function NoteBookAdd()
%>
	<table border="0" width="80%" cellpadding="0" cellspacing="0" align="center">
		<tr>
			<td align="center">
				<%Call TbTopic("ǩд����")%>
				<table border="0" cellpadding="3" cellspacing="0" class=xtable align="center">
					<form action="notebook_indb.asp" method="post" onsubmit="return checkinf(this);">
					<tr><td class=ltd>��������:</Td></Tr>
					<Tr><Td class=ltd>
							<textarea name="lylr" cols="90" rows="15"></textarea>
							<input type="hidden" name="indbinf" value="add">
					</Td></Tr>
					<Tr><Td class=ctd><input type="submit" value=" �ύ���� "></Td></Tr>
					</form>
				</table>
			</td>
		</tr>
		<Tr><Td height=10></td></tr>
	</table>
<%
End Function
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