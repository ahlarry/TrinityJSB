<!--#include file="include/conn.asp"-->
<!--#include file="js/jsCookie.js"-->
<%
Call ChkPageAble(7)
CurPage="ͼ������ �� ��Ӵ浵"
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
			<%Response.Write(XjLine(1,"100%",web_info(12)))%>
			<%Call waitSave()%>

			<%Response.Write(XjLine(10,"100%",""))%>

		</Td></Tr>
	</Table>
<%
End Sub

Function DocBak()
	Dim s_lsh
	s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("��ѡ����Ҫ���̵��������ˮ��!") : Exit Function
	strSql="select mh,lsh,ddh,dwmc,dmmc,cp from [mtask] where lsh='"&s_lsh&"' or ddh='"&s_lsh&"' and not(cp)"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("��ˮ�Ż򶩵���Ϊ ��"&s_lsh&"�� ģ�����񲻴��ڣ�","docbak_add.asp")
	ElseIf Rs("cp") Then
		Call JsAlert("��ˮ�Ż򶩵���Ϊ ��"&s_lsh&"�� ģ��ͼֽ�Ѿ����̣�","docbak_add.asp")
	Else
		Call SelectSave (rs,s_lsh)
	End IF
	Rs.Close
End Function

Function DocBakAdd(rs)
%>

	<%Call TbTopic("�����ˮ��" & rs("lsh") & " ģ��ͼֽ�浵��Ϣ")%>
	<table width="68%" cellpadding=2 cellspacing=0 class=xtable>
		<form name="frm_docbak" id="frm_docbak" action="docbak_indb.asp" method="post" onsubmit='return docbak_checkinf();'>
		<tr>
			<td class=rtd>ģ��</td>
			<td class=ltd><%=UCase(rs("mh") & "-" & rs("lsh"))%></td>
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
			<td class=ltd><input type="text" name="diskid" size="15"></td>
		</tr>
		<tr>
			<td class=rtd>��ע</td>
			<td class=ltd><textarea name="bz" cols="60" rows="10"></textarea></td>
		</tr>
		<tr>
			<td class=ctd colspan="2">
				<input type="hidden" name="lsh1" value="<%=rs("lsh")%>">
				<input type="hidden" name="ddh1" value="<%=rs("ddh")%>">
				<input type="hidden" name="mh1" value="<%=rs("mh")%>">
				<input type="hidden" name="dwmc1" value="<%=rs("dwmc")%>">
				<input type="hidden" name="lsno" value=1>
				<input type="hidden" name="indbinf" value="add">
				<input type="submit" value=" ȷ�� ">
			</td>
		</tr>
		</form>
	</table>
<%
End Function
%>
<%
Function DocBakAdds(rs)
%>

		<form name="frm_docbak" id="frm_docbak" action="docbak_indb.asp" method="post" onsubmit='return docbak_checkinf();'>

		<table width="68%" cellpadding=2 cellspacing=0 class=xtable>
		<tr>
		<td class=rtd>������</td>
			<td class=ltd><%=rs("ddh")%></td>
		</tr>
			<tr>
		<td class=rtd>��λ����</td>
			<td class=ltd><%=rs("dwmc")%></td>
		</tr>
		<tr>
			<td class=rtd>�����̺�</td>
			<td class=ltd><input type="text" name="diskid" size="15"></td>
		</tr>
		<tr>
			<td class=rtd>��ע</td>
			<td class=ltd><textarea name="bz" cols="60" rows="10"></textarea></td>
		</tr>

	<%
    Dim s_lsh,i,n
	i=1
	s_lsh=Trim(Request("s_lsh"))
	if s_lsh=rs("lsh") Then
		strsql="select *  from mtask where ddh = (select ddh from mtask where lsh='"&s_lsh&"') and  not(cp)"
	Else
		strsql="select *  from mtask where ddh='"&s_lsh&"' and  not(cp)"
	End If
	Rs.Close
	Rs.open strsql,Conn ,1,3
	if not Rs.BOF then
	Rs.MoveLast

    n=Rs.RecordCount
	end if
	Response.Write("�ö�������"&n&"����ˮ�ż�¼׼������!")
	Rs.MoveFirst
	do while not Rs.eof
'Call TbTopic("�����ˮ��" & rs("lsh") & " ģ��ͼֽ�浵��Ϣ")
%>
		<tr>
			<td class=ctd colspan="2">
			<input type="hidden" name="lsh<%=i%>" value="<%=rs("lsh")%>">
			<input type="hidden" name="ddh<%=i%>" value="<%=rs("ddh")%>">
			<input type="hidden" name="mh<%=i%>" value="<%=rs("mh")%>">
			<input type="hidden" name="dwmc<%=i%>" value="<%=rs("dwmc")%>">




<%
i=i+1
	Rs.MoveNext
	Loop
	 %>

		<input type="hidden" name="indbinf" value="add">
		<input type="hidden" name="lsno" value="<%=n%>">
		<input type="submit" value=" ȷ�� ">
		</td>
		</tr>

	  </table>
	  </form>

<%
	End Function
%>
<%
Function waitSave()
	Dim strAll
	strAll=request("disall")
	If strAll="" Then strAll="yes"
%>

	<%Call TbTopic("�ȴ��浵��ģ��")%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable>
		<tr>
			<th class=th>id</th>
			<th class=th>������</th>
			<th class=th>��ˮ��</th>
			<th class=th>��λ����</th>
			<th class=th>��������</th>
			<th class=th>���ʱ��</th>
			<th class=th>ģ��</th>
			<%if ChkAble(7) then response.write("<td class=th>����</td>")%>
		</tr>
		<%
			Dim i
			Set Rs = xjweb.Exec("select * from [mtask] where not(isnull(sjjssj)) and not(cp) order by id",1)
			i = 1
			do while not rs.eof
		%>
				<tr>
					<td class=ctd><%=i%></td>
					<td class=ctd><%=rs("ddh")%></td>
					<td class=ctd><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
					<td class=ctd><%=rs("dwmc")%></td>
					<td class=ctd><%=rs("dmmc")%></td>
					<td class=ctd><%=rs("sjjssj")%></td>
					<td class=ctd><%=ucase(rs("mh") & "-" &  rs("lsh"))%>
					</td>

	                <% response.write("<td class=ctd><a href=""docbak_add.asp?s_lsh="&rs("lsh")&""" onclick=""getUserSelect() ;"">����</a></td>")%>
				</tr>
			<%
				if i >= 20 and strAll <> "yes" then exit do
				i=i+1
				rs.movenext
			loop
			rs.close
			set rs = nothing
		%>

	</table><br>
	<%if strAll <> "yes" then response.write("<a href='?disall=yes'>��ʾ����</a>")%>


<%end function%>
<%function SelectSave (rs,s_lsh)
dim temp
temp =Request.Cookies("useroperation")
if rs("ddh")=s_lsh Then temp="batch"
if temp="batch" then
Call DocBakAdds(rs)
else
Call DocBakAdd(rs)
end if
end function
 %>