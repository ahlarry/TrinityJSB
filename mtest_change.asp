<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(6)
CurPage="ģ�ߵ��� �� ���ĵ�����Ϣ"					'ҳ�������λ��( ��������� �� ���������)
strPage="mtest"
Call FileInc(0, "js/mtest.js")
xjweb.header()
Call TopTable()

Dim iid, strlsh, ics, bps
iid=clng(Request("id"))
strlsh=Request("s_lsh")
ics=clng(Request("cs"))
bps=Request("ps")

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchLsh()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call mtestChange()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function mtestChange()
	If strlsh="" Then Call JsAlert("����������!","mtest_list.asp") : Exit Function
	strSql="select * from [ts_tsxx] where id=" & iid
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("������!����������!","mtest_list.asp") : Exit Function
	Else
		Response.Write(xjLine(10, "100%", ""))
		If bps Then
			Call mtestps_change(Rs)
		Else
			Call mtest_change(Rs)
		End If
	End If
	Rs.close
End Function


Function mtest_change(Rs)
%>
	<%Call TbTopic("������ˮ�� "&strlsh&" ģ�ߵ� " & ics &" �ε�����Ϣ")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%">
	<form id=frm_mtestadd name=frm_mtestadd action=mtest_indb.asp?action=change method=post onSubmit='return tscheckinf();'>

	<tr>
		<th class=rtd height=25 width="20%">��Ŀ����</td>
		<th class=ctd width="*">��Ŀ����</td>
	</tr>
	<tr>
		<td class=rtd>����ԭ��</td>
		<td class=ltd><textarea name="tsyy" cols="95" rows="7"><%=Rs("tsyy")%></textarea></td>
	</tr>

	<tr>
		<td class=rtd>��������</td>
		<td class=ltd><textarea name="tslr" cols="95" rows="7"><%=Rs("tslr")%></textarea></td>
	</tr>
	<tr><td class=ctd colspan=2><input type=submit value=" �� �� �� �� "></td></tr>
	<input type="hidden" name="id" value=<%=iid%>>
	<input type="hidden" name="lsh" value=<%=Rs("lsh")%>>
	</form>

	</table>
<%
End Function		'mtest_change()

Function mtestps_change(Rs)
%>
	<%Call TbTopic("�����ˮ�� " & strlsh & " ģ�ߵ� " & ics & " �������¼") %>
	<table class=xtable cellspacing=0 cellpadding=3 width="98%">
	<form id=frm_mtestpsadd name=frm_mtestpsadd action=mtest_indb.asp?action=change method=post onSubmit='return tspscheckinf();'>

	<tr>
		<th class=rtd height=25 width="20%">��Ŀ����</td>
		<th class=ctd width="*">��Ŀ����</td>
	</tr>
	<tr>
		<td class=rtd>��������</td>
		<td class=ltd><textarea name="tslr" cols="95" rows="7"><%=Rs("tslr")%></textarea></td>
	</tr>

	<tr>
		<td class=rtd>������</td>
		<td class=ltd><textarea name="tsyy" cols="95" rows="3"><%=Rs("tsyy")%></textarea></td>
	</tr>
	<tr><td class=ctd colspan=2><input type=submit value=" �� �� �� �� "></td></tr>
	<input type="hidden" name="id" value=<%=iid%>>
	<input type="hidden" name="lsh" value=<%=Rs("lsh")%>>
	</form>

	</table>
<%
End Function
%>