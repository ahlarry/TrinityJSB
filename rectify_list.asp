<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="������� �� ����/Ԥ����ʩ��"
strPage="tech"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Dim strFeedBack,strbh,strzrbm,strxxbm,strjssj,strbhgnr,strwczk
strbh=Request("bh")
strzrbm=Request("zrbm")
strxxbm=Request("xxbm")
strjssj=Request("jssj")
strbhgnr=Request("bhgnr")
strwczk=Request("wczk")
strFeedBack="&zrbm="&strzrbm&"&xxbm="&strxxbm&"&bhgnr="&strbhgnr

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
			<%Call RectifyList()%>
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
			���:<input type="text" name="bh" size=8 value="<%=strbh%>">
			���β���:<input type="text" name="zrbm" size=8 value="<%=strzrbm%>">
			���ϸ�/Ǳ�ڲ��ϸ�����:<input type="text" name="bhgnr" size=30 value="<%=strbhgnr%>">
		<input type=submit value=" ѡ �� ">
		</Td></Tr>
		</Form>
	</Table>
<%
End Function

Function RectifyList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount
	absPageNum = 0
	RecordPerPage = 20
	strSql=""
	If strbh<>"" Then strSql=" bh='" & strbh & "'" End If
	If strzrbm<>"" Then
		strSql=strSql & " zrbm='"&strzrbm&"' "
	End If
	If strbhgnr<>"" Then
		If strSql<>"" Then
			strSql=strSql & " and bhgnr like '%" & strbhgnr &"%'"
		Else
			strSql=" bhgnr like '%" & strbhgnr &"%'"
		End IF
	End If

	If strSql<>"" Then
		strSql="select * from [Rectify]  where" & strSql & " order by jssj desc"
	Else
		strSql="select * from [Rectify] order by jssj desc"
	End If

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
	Call TbTopic("����/Ԥ����ʩ��")

%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable  align="center">
		<tr>
			<th class=th width=40>ID</th>
			<th class=th width=40>���</th>
			<th class=th width=60>���β���</th>
			<th class=th width=80>������Ϣ����</th>
			<th class=th width=*>���ϸ�/Ǳ�ڲ��ϸ�����</th>
			<th class=th width=100>��Ϣ��������</th>
			<th class=th width=60>��ǰ״̬</th>
		</tr>
<%
	for absrecordnum = 1 to recordperpage
%>

		<tr>
			<td class=ctd><%=icounter%></td>
			<td class=ctd alt="����鿴������Ϣ"><a href="Rectify_dis.asp?s_lsh=<%=rs("bh")%>&id=<%=rs("id")%>"><%=rs("bh")%></a></td>
			<td class=ctd><%=rs("zrbm")%></td>
			<td class=ctd><%=rs("xxbm")%></td>
			<td class=ltd><%= xjweb.StringCut(rs("bhgnr"),26)%></td>
			<td class=ctd><%=rs("jssj")%></td>
			<td class=ctd><%=rs("wczk")%></td>
		</tr>
		<%rs.movenext%>
		<%if rs.eof then%>
			<%exit for%>
		<%end if%>
		<%icounter = icounter - 1%>
	<%next%>
	</table>
	<table width="95%" cellpadding=2 cellspacing=0 border=0 align="center">
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