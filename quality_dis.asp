<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="������� �� �鿴�ⲿ������Ϣ"
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
			<%Call qualityDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function qualityDisplay()
	Dim s_hth, iid, strAlt
	s_hth="" : iid=0 : strAlt=""
	s_hth=Trim(Request("s_lsh"))
	If IsNumeric(Trim(Request("id"))) Then iid=CInt(Trim(Request("id")))
	If s_hth="" Then Call TbTopic("������鿴�ⲿ������Ϣ�ĺ�ͬ��!") : Exit Function
	strSql="select lxr, lxdh, gzlh,jssj from [quality] where hth='"&s_hth&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Not(Rs.Eof Or Rs.Bof) Then
		strAlt="��ϵ��:" & Rs("lxr") & "<br>��ϵ�绰:" & Rs("lxdh") & "<br>�������" & Rs("gzlh") & "<br>����ʱ��:" & Rs("jssj")
	End If
	Rs.Close

	strSql="select * from [quality] where hth='"&s_hth&"'"
	If iid<>0 Then strSql=strSql & " and id="&iid&""
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("��ͬ�� ��" & s_lsh & "�� û���κ�����! ","quality_list.asp")
	Else
		Call quality_Display(Rs, strAlt)
	End If
	Rs.close
End Function

Function Task_Info(Rs)
%>
	<%Call TbTopic("��ͬ�� "&Rs("hth")&" ������Ϣ")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
		<tr>
			<td class=th width="20%">��ϵ��</td>
			<td class=th width="*">��ϵ�绰</td>
			<td class=th width="20%">�������</td>
			<td class=th width="20%">����ʱ��</td>
		</tr>

		<tr>
			<td class=ctd><%=Rs("lxr")%></a></td>
			<td class=ctd><%=Rs("lxdh")%></td>
			<td class=ctd><%=Rs("gzlh")%></td>
			<td class=ctd><%=Rs("jssj")%></td>
		</tr>
	</table>
<%
End Function

Function quality_Display(Rs, strAlt)
	Call TbTopic("��ͬ�� " &Rs("hth")&" �ⲿ������Ϣ")
%>
	<%Do While Not Rs.Eof%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
		<Tr>
			<Td class=th width=80>��ͬ��</Td>
			<Td class=ctd width=80 <%If strAlt<>"" Then%>alt="<%=strAlt%>"<%End If%>><%=Rs("hth")%></Td>
			<Td class=th width=80>�ͻ�����</Td>
			<Td class=ctd width=*><%=Rs("khmc")%></Td>
			<Td class=th width=80>������</Td>
			<Td class=ctd width=80><%=Rs("zrr")%></Td>
			<Td class=th width=80>��ǰ״̬</Td>
			<Td class=ctd width=80><%=Rs("wczk")%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>��Ҫ����</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("zywt"))%></Td>
		</Tr>

		<% If Rs("yjcs") <> "" Then %>
		<Tr>
			<Td class=th width=80>Ӧ����ʩ</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("yjcs"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("yyfx") <> "" Then %>
		<Tr>
			<Td class=th width=80>ԭ�����</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("yyfx"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("jzcs") <> "" Then %>
		<Tr>
			<Td class=th width=80>������ʩ</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("jzcs"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("lsqk") <> "" Then %>
		<Tr>
			<Td class=th width=80>��ʵ���</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("lsqk"))%></Td>
		</Tr>
		<% End If %>

		<% If Rs("yzjl") <> "" Then %>
		<Tr>
			<Td class=th width=80>��֤����</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("yzjl"))%></Td>
		</Tr>
		<% End If %>

		<%If ChkAble(11) Then%>
		<Tr>
			<Td colspan=4 class=ctd><a href="quality_change.asp?id=<%=Rs("id")%>">����</a></Td>
			<Td colspan=4 class=ctd><a href="quality_indb.asp?action=delete&id=<%=Rs("id")%>" onclick="return confirm('ȷ��ɾ����?');">ɾ��</a></Td>
		</Tr>
		<%End If%>
		</Table>
	<%
			Response.write(XjLine(10, "100%", ""))
			Rs.MoveNext
		Loop
End Function
%>