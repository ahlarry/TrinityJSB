<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(7)
CurPage="������� �� ����������"
strPage="tech"
Call FileInc(0, "js/tech.js")
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
			<%Call techAdd()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

function techAdd()
%>
	<%Call TbTopic("��Ӽ����������")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<form id=frm_techadd name=frm_techadd action=tech_indb.asp?action=add method=post onSubmit='return checkinf();'>

	<tr>
		<th class=th height=25>��Ŀ����</td>
		<th class=th>��Ŀ����</td>
	</tr>

	<tr>
		<td class=rtd>��ˮ��</td>
		<td class=ltd><input type=text name="lsh" size=15></td>
	</tr>
	<tr>
		<td class=rtd>�������</td>
		<td class=ltd><input type=text name="bkmc" size=15></td>
	</tr>

	<tr>
		<td class=rtd width="20%">�������</td>
		<td class=ltd>
			<select name="clyj">
				<option value="����">����</option>
				<option value="����">����</option>
				<option value="����">����</option>
			</select>
		</td>
	</tr>

	<tr>
		<td class=rtd>������</td>
		<td class=ltd><input type=text name="zrr" size=15></td>
	</tr>

	<tr>
		<td class=rtd valign=top>������������</td>
		<td class=ltd><textarea name="xxms" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd valign=top>����ԭ�����</td>
		<td class=ltd><textarea name="yyfx" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd valign=top>����Ԥ����ʩ</td>
		<td class=ltd><textarea name="yfcs" cols="75" rows="7"></textarea></td>
	</tr>

	<tr><td class=ctd colspan=2><input type=submit value=" �� ȷ �� �� "></td></tr>
	</form>
	</table>
<%
end function		'tech_add()
%>