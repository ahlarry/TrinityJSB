<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(7)
CurPage="������� �� �����������"
strPage="tech"
Call FileInc(0, "js/tech.js")
xjweb.header()
Call TopTable()
Dim iid
iid=0
If IsNumeric(Trim(Request("id"))) Then iid=CLng(Trim(Request("id")))
If iid=0 Then
	Call JsAlert("���������ڽ���!лл!","tech_list.asp")
Else
	Call Main()
End If
Call BottomTable()
xjweb.footer()
closeObj()

Function Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%
				If iid=0 Then Exit Function
				strSql="select * from [tecq_question] where id="&iid&""
				Set Rs=xjweb.Exec(strSql, 1)
				If Rs.Eof Or Rs.Bof Then Rs.Close : Exit Function
				Call techChange(Rs)
				Rs.Close
				Response.Write(XjLine(10,"100%",""))
			%>
		</Td></Tr>
	</Table>
<%
End Function

Function techChange(Rs)
	Call TbTopic("���ļ����������")
%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<form id=frm_techadd name=frm_techadd action=tech_indb.asp?action=change method=post onSubmit='return checkinf();'>

	<tr>
		<th class=th height=25>��Ŀ����</td>
		<th class=th>��Ŀ����</td>
	</tr>

	<tr>
		<td class=rtd>��ˮ��</td>
		<td class=ltd><input type=text name="lsh" size=15 value="<%=Rs("tecq_lsh")%>"></td>
	</tr>
	<tr>
		<td class=rtd>�������</td>
		<td class=ltd><input type=text name="bkmc" size=15 value="<%=Rs("tecq_bkmc")%>"></td>
	</tr>

	<tr>
		<td class=rtd width="20%">�������</td>
		<td class=ltd>
			<select name="clyj">
				<option value="<%=Rs("tecq_clyj")%>"><%=Rs("tecq_clyj")%></option>
				<option value="����">����</option>
				<option value="����">����</option>
			</select>
		</td>
	</tr>

	<tr>
		<td class=rtd>������</td>
		<td class=ltd>
			<select name="zrr"><option></option>
				<%for i = 0 to ubound(c_jsb)%>
					<option value="<%=c_jsb(i)%>" <%If Rs("tecq_zrr")=c_jsb(i) Then%> Selected<%End If%>><%=c_jsb(i)%></option>
				<%next%>
			</select>
		</td>
	</tr>

	<tr>
		<td class=rtd valign=top>������������</td>
		<td class=ltd><textarea name="xxms" cols="75" rows="7"><%=Rs("tecq_xxms")%></textarea></td>
	</tr>

	<tr>
		<td class=rtd valign=top>����ԭ�����</td>
		<td class=ltd><textarea name="yyfx" cols="75" rows="7"><%=Rs("tecq_yyfx")%></textarea></td>
	</tr>

	<tr>
		<td class=rtd valign=top>����Ԥ����ʩ</td>
		<td class=ltd><textarea name="yfcs" cols="75" rows="7"><%=Rs("tecq_yfcs")%></textarea></td>
	</tr>

	<tr><td class=ctd colspan=2><input type=submit value=" �� ȷ �� �� "></td></tr>
	<input type="hidden" name="id" value=<%=Rs("id")%>>
	</form>
	</table>
<%
end function
%>