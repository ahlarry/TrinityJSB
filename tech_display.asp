<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
CurPage="������� �� �鿴�������"					'ҳ�������λ��( ��������� �� ���������)
strPage="tech"
'Call FileInc(0, "js/mtest.js")
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
			<%Call techDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function techDisplay()
	Dim s_lsh, iid, strAlt
	s_lsh="" : iid=0 : strAlt=""
	s_lsh=Trim(Request("s_lsh"))
	If IsNumeric(Trim(Request("id"))) Then iid=CInt(Trim(Request("id")))
	If s_lsh="" Then Call TbTopic("������鿴���������ģ����ˮ��!") : Exit Function
	strSql="select lsh, dmmc, mh,dwmc from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Not(Rs.Eof Or Rs.Bof) Then
		strAlt="��ˮ��:" & Rs("lsh") & "<br>��λ����:" & Rs("dwmc") & "<br>��������" & Rs("dmmc") & "<br>ģ��:" & Rs("mh")
	End If
	Rs.Close

	strSql="select * from [tecq_question] where tecq_lsh='"&s_lsh&"'"
	If iid<>0 Then strSql=strSql & " and id="&iid&""
	'Response.Write strSql
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("��ˮ�� ��" & s_lsh & "�� ������û���κ�����! ","tech_list.asp")
	Else
		Call tech_Display(Rs, strAlt)
	End If
	Rs.close
End Function

Function Task_Info(Rs)
%>
	<%Call TbTopic("��ˮ�� "&Rs("lsh")&" ģ����Ϣ")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%">
		<tr> 
			<td class=th width="20%">��ˮ��</td>
			<td class=th width="*">��������</td>
			<td class=th width="20%">ģ��</td>
			<td class=th width="20%">��λ����</td>
		</tr>
								
		<tr> 
			<td class=ctd><a href="mtask_display.asp?s_lsh=<%=Rs("lsh")%>" alt="�鿴��ˮ�� <b><%=Rs("lsh")%></b> ������"><%=Rs("lsh")%></a></td>
			<td class=ctd><%=Rs("dmmc")%></td>
			<td class=ctd><%=Rs("mh")%></td>
			<td class=ctd><%=Rs("dwmc")%></td>
		</tr>
	</table>	
<%
End Function

Function tech_Display(Rs, strAlt)
	Call TbTopic("��ˮ�� " &Rs("tecq_lsh")&" ģ�߼����������")
%>
	<%Do While Not Rs.Eof%>
	<Table class=xtable cellspacing=0 cellpadding=3 width="95%">
		<Tr>
			<Td class=th width=80>��ˮ��</Td>
			<Td class=ctd width=80 <%If strAlt<>"" Then%>alt="<%=strAlt%>"<%End If%>><a href="mtask_display.asp?s_lsh=<%=Rs("tecq_lsh")%>" alt="�鿴��ˮ�� <b><%=Rs("tecq_lsh")%></b> ������"><%=Rs("tecq_lsh")%></a></Td>
			<Td class=th width=80>�������</Td>
			<Td class=ctd width=*><%=Rs("tecq_bkmc")%></Td>
			<Td class=th width=80>ģ��</Td>
			<Td class=ctd width=80><%=Rs("tecq_clyj")%></Td>
			<Td class=th width=80>������</Td>
			<Td class=ctd width=80><%=Rs("tecq_zrr")%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>������������</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("tecq_xxms"))%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>����ԭ�����</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("tecq_yyfx"))%></Td>
		</Tr>
		<Tr>
			<Td class=th width=80>����Ԥ����ʩ</Td>
			<Td class=ltd colspan=7><%=xjweb.htmltocode(Rs("tecq_yfcs"))%></Td>
		</Tr>
		<%If ChkAble("1,7") Then%>
		<Tr>
			<Td colspan=4 class=ctd><a href="tech_change.asp?id=<%=Rs("id")%>">����</a></Td>
			<Td colspan=4 class=ctd><a href="tech_indb.asp?action=delete&id=<%=Rs("id")%>" onclick="return confirm('ȷ��ɾ����?');">ɾ��</a></Td>
		</Tr>
		<%End If%>
		</Table>
	<%	
			Response.write(XjLine(10, "100%", ""))
			Rs.MoveNext
		Loop
	%>
	
	
<%
End Function
%>