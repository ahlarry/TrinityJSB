<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(7)
CurPage="ͼ������ �� ���Ĵ浵"
strPage="docbak"
Call FileInc(0, "js/docbak.js")
xjweb.header()
Call TopTable()
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
			<%Call DocBak()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function DocBak()
	Dim s_lsh
	s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("��������Ĵ���ģ�ߵ���ˮ�ţ�") : Exit Function
	strSql="select a.mh,a.lsh,a.ddh,a.dwmc,a.dmmc, b.diskid,b.bz from [mtask] a, [doc_bak] b where b.lsh='"&s_lsh&"' and a.lsh = b.lsh"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("��ˮ�� ��"&s_lsh&"�� ģ�����񲻴��ڣ�\n����ԭ����: 2004����ǰ��ģ�߻���û�д浵��","docbak_change.asp")
	Else
		Call DocBakChange(rs)
	End If
	Rs.Close
End Function

Function DocBakChange(rs)
%>
	<%Call TbTopic("������ˮ�� "&rs("lsh")&" ͼֽ�浵��Ϣ")%>
	<table cellpadding=2 cellspacing=0 class=xtable>
		<form name="frm_docbak" id="frm_docbak" action="docbak_indb.asp" method="post" onsubmit='return docbak_checkinf();'>
		<tr>
			<td class=rtd>ģ��</td>
			<td class=ltd><%=ucase(rs("mh") & "-" & rs("lsh"))%></td>
		</tr>
		<tr>
			<td class=rtd>��λ����</td>
			<td class=ltd><%=rs("dwmc")%></td>
		</tr>
		<tr>
			<td class=rtd>��������</td>
			<td class=ltd><%=rs("dmmc")%></td>
		</tr>
		<tr>
			<td class=rtd>�����̺�</td>
			<td class=ltd><input type="text" name="diskid" size="15" value="<%=rs("diskid")%>"></td>
		</tr>
		<tr>
			<td class=rtd>��ע</td>
			<td class=ltd><textarea name="bz" cols="60" rows="10"><%=rs("bz")%></textarea></td>
		</tr>
		<tr>
			<td class=ctd colspan="2">
				<input type="hidden" name="lsh" value="<%=rs("lsh")%>">
				<input type="hidden" name="ddh" value="<%=rs("ddh")%>">
				<input type="hidden" name="mh" value="<%=rs("mh")%>">
				<input type="hidden" name="dwmc" value="<%=rs("dwmc")%>">
				<input type="hidden" name="indbinf" value="change">
				<input type="submit" value=" �� �� ">
			</td>
		</tr>
		</form>
	</table>
<%
End Function
%>