<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="������� �� �鿴����/Ԥ����ʩ��"
strPage="tech"
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
			<%Call RectifyDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function RectifyDisplay()
	Dim s_hth, iid
	s_hth="" : iid=0
	s_hth=Trim(Request("s_lsh"))
	If IsNumeric(Trim(Request("id"))) Then iid=CInt(Trim(Request("id")))
	If s_hth="" Then Call TbTopic("������鿴����/Ԥ����ʩ��ı��!") : Exit Function

	strSql="select * from [Rectify] where bh='"&s_hth&"'"
	If iid<>0 Then strSql=strSql & " and id="&iid&""
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("��� ��" & s_lsh & "�� ����/Ԥ����ʩ������! ","Rectify_list.asp")
	Else
		Call Rectify_Display(Rs)
	End If
	Rs.close
End Function

Function Rectify_Display(Rs)
	Call TbTopic("��� " &Rs("bh")&" ����/Ԥ����ʩ��")
%>
	<%Do While Not Rs.Eof%>
		<Table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
		<tr>
			<td class=th width="20%">���β���</td>
			<Td class=ctd><%=Rs("zrbm")%></Td>
			<td class=th width="20%" colspan="2">���</td>
			<Td class=ctd><%=Rs("bh")%></Td>
		</tr>

		<tr>
			<td class=th>������Ϣ����</td>
			<Td class=ctd><%=Rs("xxbm")%></Td>
			<td class=th colspan="2">��Ϣ��������</td>
			<Td class=ctd><%=Rs("jssj")%></Td>
		</tr>

		<tr>
			<td class=th>���ϸ�/Ǳ�ڲ��ϸ�����</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("bhgnr"))%></td>
		</tr>

		<% If Rs("yfcsyq") <> "" Then %>
		<tr>
			<td class=th rowspan="2">����/Ԥ����ʩҪ��</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("yfcsyq"))%></td>
		</tr>

		<tr>
			<td class=ctd colspan="2">����:<%=Rs("qxsj")%></td>
			<td class=ctd colspan="2">�Ƿ�����<% if Rs("ps")="V1" Then %>��<% else %>��<% End if %></td>
		</tr>
		<% End If %>
		<% If Rs("yyfx") <> "" Then %>
		<tr>
			<td class=th>ԭ�����</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("yyfx"))%></td>
		</tr>
		<% End If %>
		<% If Rs("jzcs") <> "" Then %>
		<tr>
			<td class=th>������ʩ</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("jzcs"))%></td>
		</tr>
		<% End If %>
		<% If Rs("lsqk") <> "" Then %>
		<tr>
			<td class=th>��ʵ���</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("lsqk"))%></td>
		</tr>
		<% End If %>
		<% If Rs("yzjl") <> "" Then %>
		<tr>
			<td class=th>��֤����</td>
			<td class=ctd colspan="4"><%=xjweb.htmltocode(Rs("yzjl"))%></td>
		</tr>
		<% End If %>
		<%If ChkAble(11) Then%>
		<tr>
			<td class=ctd colspan="2" align="center"><a href="rectify_change.asp?id=<%=Rs("id")%>">����</a></td>
			<td class=ctd colspan="3" align="center"><a href="rectify_indb.asp?action=delete&id=<%=Rs("id")%>" onclick="return confirm('ȷ��ɾ����?');">ɾ��</a></td>
		</tr>
		<% End If %>
	</table>
	<%
			Response.write(XjLine(10, "100%", ""))
			Rs.MoveNext
		Loop
End Function
%>