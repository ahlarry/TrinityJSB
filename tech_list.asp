<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="������� �� ��������б�"
strPage="tech"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

Dim strFeedBack,strZrr, strLsh, strXx
strZrr=Trim(Request("zrr"))
strLsh=Request("lsh")
strXx=Trim(Request("wtxx"))
strFeedBack="&lsh="&strLsh&"&wtxx="&strXx&"&zrr="&strZrr

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchInfo()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call techList()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function SearchInfo()
%>
	<Table border=0 cellpadding=2 cellspacing=0 width="100%">
		<Form action="<%=Request.Servervariables("script_name")%>" method=get>
		<Tr><Td>
		&nbsp;&nbsp;ѡ��ɸѡ����: 
			��ˮ��:<input type="text" name="lsh" size=8 value="<%=strLsh%>">
			��������:<input type="text" name="wtxx" size=30 value="<%=strXx%>">
			������:
			<Select name="zrr">
				<option value="">ȫ��</option>
				<%If Session("userName")<>"" Then%>
					<option value="<%=Session("userName")%>"><%=Session("userName")%></option>
				<%End If%>
				<%If ChkAble("1,2,3") Then%>
					<%for i = 0 to ubound(c_allzy)%>
						<option value="<%=c_allzy(i)%>"<%If strZrr=c_allzy(i) Then%> Selected<%End If%>><%=c_allzy(i)%></option>
					<%next%>
				<%End If%>
			</Select>
		<input type=submit value=" ѡ �� ">
		</Td></Tr>
		</Form>
	</Table>
<%
End Function

Function techList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount
	absPageNum = 0
	RecordPerPage = 20
	strSql=""
	If strLsh<>"" Then strSql=" tecq_lsh='" & strLsh & "'"
	If strXx<>"" Then
		If strSql<>"" Then
			strSql=strSql & " and tecq_xxms like '%" & strXx &"%'"
		Else
			strSql=" tecq_xxms like '%" & strXx &"%'"
		End IF
	End If

	If strZrr<>"" Then
		If strSql<>"" Then
			strSql=strSql & " tecq_zrr='"&strZrr&"' "
		Else
			strSql=" tecq_zrr='"&strZrr&"' "
		End If
	End If
	If strSql<>"" Then
		strSql="select * from [tecq_question]  where" & strSql & " order by tecq_time desc"
	Else
		strSql="select * from [tecq_question] order by tecq_time desc"
	End If

	'Response.write strSql
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,3,3
	If Rs.eof Or Rs.bof Then
		Call TbTopic("û���κ����ݷ���������") : Exit Function
	End If
	Rs.PageSize = RecordPerPage
	TotalCount=Rs.RecordCount
	
	If Trim(Request("iPage")) <> ""  Then
		If IsNumeric(Trim(Request("iPage"))) Then
			If Trim(Request("iPage")) <= 0 Then
				absPageNum = 1
			ElseIf CLng(Trim(Request("iPage"))) > Rs.PageCount Then
				absPageNum = Rs.PageCount
			Else
				absPageNum = CLng(Trim(Request("iPage")))
			End If
		Else
			If Request("iCurPage") <> "" Then
				absPageNum = CLng(Request("iCurPage"))
			Else
				absPageNum = 1
			End If
		End If
	Else
		If Request("iCurPage") <> "" Then
			absPageNum = CLng(Request("iCurPage"))
		Else
			absPageNum = 1
		End If
	End If
	
	If absPageNum > Rs.PageCount then absPageNum = Rs.PageCount
	rs.absolutepage = absPageNum
	icounter=totalcount-(abspagenum-1)*recordperpage
	Call TbTopic("����ģ�߳������������")
	'icounter = (absPageNum - 1) * recordperpage + 1
%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable>
		<tr>
			<th class=th width=40>ID</th>
			<th class=th width=60>��ˮ��</th>
			<th class=th width=80>�������</th>
			<th class=th width=*>��������</th>
			<th class=th width=60>�������</th>
			<th class=th width=60>������</th>
			<th class=th width=100>����ʱ��</th>
		</tr>
<%	
	for absrecordnum = 1 to recordperpage
%>
		
		<tr>
			<td class=ctd><%=icounter%></td>
			<td class=ctd alt="����鿴������Ϣ"><a href="tech_display.asp?s_lsh=<%=rs("tecq_lsh")%>&id=<%=rs("id")%>"><%=rs("tecq_lsh")%></a></td>
			<td class=ctd><%=rs("tecq_bkmc")%></td>
			<td class=ltd><%= xjweb.StringCut(rs("tecq_xxms"),26)%></td>
			<td class=ctd><%=rs("tecq_clyj")%></td>
			<td class=ctd><%=rs("tecq_zrr")%></td>
			<td class=ctd><%=xjDate(rs("tecq_time"),1)%></td>
		</tr>
		<%rs.movenext%>
		<%if rs.eof then%>
			<%exit for%>
		<%end if%>
		<%icounter = icounter - 1%>
	<%next%>
	</table>
	<table width="95%" cellpadding=2 cellspacing=0 border=0>
		<tr>
			<td align=left>
				���������� <%=rs.recordcount%> ��&nbsp;&nbsp;
				ÿҳ <%=rs.pagesize%> ��&nbsp;&nbsp;
				�� <%=Rs.PageCount%> ҳ&nbsp;&nbsp;
				��ǰΪ�� <%=absPageNum%> ҳ
			</td>
			<td align=right>
				��
			<%
				if absPageNum > 1 then
					response.write("<a href="&Request.ServerVariables("script_name")&"?ipage="&(abspagenum-1)&strFeedBack&" alt='��һҳ'> ��</a>&nbsp;&nbsp;")
				end if
				Dim iStart,iEnd
				if absPageNum < 4 then
					iStart = 1
				else
					iStart = absPageNum - 3
				end if
				if absPageNum < Rs.PageCount - 3 then
					iEnd = absPageNum + 3
				else
					iEnd = Rs.PageCount
				end if
				for i = iStart to iEnd
					if i = absPageNum then
						response.write("&nbsp;<font style=font-size:11pt;><b>"&  i & "</b></font>&nbsp;")
					else
						response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&i&strFeedBack&">" & i & "</a>&nbsp;")
					end if
				next
				if absPageNum < Rs.PageCount then
					response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&(absPageNum+1)&strFeedBack&" alt='��һҳ'> �� </a>&nbsp;")
				end if
			%>	
				��
				��ת��:
				<select name="ipage" onchange='location.href("<%=Request.ServerVariables("script_name")%>?ipage=" + this.value+"<%=strFeedBack%>");'>
					<%for i=1 to Rs.PageCount%>
						<%if i = absPageNum then%>
							<option value=<%=i%> selected>�� <%=i%> ҳ</option>
						<%else%>
							<option value=<%=i%>>�� <%=i%> ҳ</option>
						<%end if%>
					<%next%>
				</select>
				</td>
			</tr>
		</table>
<%
end function
%>