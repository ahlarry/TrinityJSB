<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="Ա������ �� Ա�������б�"
strPage="ygkp"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

Dim strFeedBack,strZrr, strkpitem, strgzz, strkpjs, strclsh, striPage
striPage=NullToNum(Request("ipage"))
strZrr=Trim(Request("zrr"))
strkpitem = trim(request("kpitem"))
strkpjs = trim(request("kpjs"))
strclsh = trim(request("clsh"))
strgzz =request("gzz")
If strgzz="" Then strgzz=0
strFeedBack=""
If strZrr<>"" Then strFeedBack="&zrr="&strZrr
If strkpitem<>"" Then strFeedBack="&kpitem="&strkpitem&strFeedBack
If strgzz<>"0" Then strFeedBack="&gzz="&strgzz&strFeedBack
If strkpjs<>"" Then strFeedBack="&kpjs="&strkpjs&strFeedBack
If strclsh<>"" Then strFeedBack="clsh="&strclsh&"&"&strFeedBack

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
			<%Call ygkpList()%>
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
		&nbsp;&nbsp;
			������:
			<Select name="zrr" onchange="javascript:window.location.href('<%=request.servervariables("script_name")%>?zrr=' + this.value+'&kpitem='+this.form.kpitem.value+'&kpjs='+this.form.kpjs.value);">
				<option value="">ȫ��</option>
				<%If Session("userName")<>"" Then%>
					<option value="<%=Session("userName")%>"><%=Session("userName")%></option>
				<%End If%>
				<%If ChkAble("1,2,3,4,10,11") Then%>
					<%for i = 0 to ubound(c_allzy)%>
						<option value="<%=c_allzy(i)%>"<%If strZrr=c_allzy(i) Then%> Selected<%End If%>><%=c_allzy(i)%></option>
					<%next%>
				<%End If%>
			</Select>&nbsp;&nbsp;&nbsp;
			�����飺
				<Select name="gzz" onchange="javascript:window.location.href('<%=request.servervariables("script_name")%>?gzz=' + this.value+'&kpitem='+this.form.kpitem.value+'&kpjs='+this.form.kpjs.value);">
				<option value=0>ȫ��</option>
				<%If ChkAble("1,2,3,4") Then%>
					<%for i = 1 to 5%>
						<option value=<%=i%> <%If Cint(strgzz)=i Then%> Selected<%End If%>><%=i%></option>
					<%next%>
				<%End If%>
			</Select>&nbsp;&nbsp;&nbsp;
			������Ŀ:&nbsp;<input type="text" name="kpitem" value="<%=strkpitem%>" size="10">&nbsp;&nbsp;
			��ɫ:&nbsp;<input type="text" name="kpjs" value="<%=strkpjs%>" size="8">&nbsp;&nbsp;
			��ˮ��:&nbsp;<input type="text" name="clsh" value="<%=strclsh%>" size="8">&nbsp;&nbsp;
		<input type=submit value=" ѡ �� ">
		</Td></Tr>
		</Form>
	</Table>
<%
End Function

Function ygkpList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount
	absPageNum = 0
	RecordPerPage = 20
	strSql=""
	if strkpitem <> "" then
			strSql = " kp_item like '%"&strkpitem&"%'"
	end if
	If strkpjs<>"" Then
		If strSql<>"" Then
			strSql=" kp_zrrjs='"&strkpjs&"' and " & strSql
		Else
			strSql=" kp_zrrjs='"&strkpjs&"' "
		End If
	End If
	If strgzz<>0 Then
		If strSql<>"" Then
			strSql=" kp_group="&cstr(strgzz)&" and " & strSql
		Else
			strSql=" kp_group="&cstr(strgzz)
		End If
	End If
	If strZrr<>"" Then
		If strSql<>"" Then
			strSql=" kp_zrr='"&strZrr&"' and " & strSql
		Else
			strSql=" kp_zrr='"&strZrr&"' "
		End If
	End If
	If strclsh<>"" Then
		If strSql<>"" Then
			strSql=" kp_lsh='"&strclsh&" or kp_bz like '%"&strclsh&"%' and " & strSql
		Else
			strSql=" kp_lsh='"&strclsh&"' or kp_bz like '%"&strclsh&"%'"
		End If
	End If
	If strSql<>"" Then
		strSql="select * from [kp_jsb]  where" & strSql & " order by kp_time desc"
	Else
		strSql="select * from [kp_jsb] order by kp_time desc"
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
	Call TbTopic("����ģ�߳�Ա�������б�")
	'icounter = (absPageNum - 1) * recordperpage + 1
%>
	<table width="95%" cellpadding=2 cellspacing=0 class=xtable  align="center">
		<tr>
			<th class=th width=40>ID</th>
			<th class=th width=*>������Ŀ</th>
			<th class=th width=60>������</th>
			<th class=th width=60>����</th>
			<th class=th width=60>������</th>
			<th class=th width=80>��ɫ</th>
			<th class=th width=100>����ʱ��</th>
			<%If ChkAble("1,2,3,11,12") Then%>
				<th class=th width=100>����</th>
			<%End If%>
		</tr>
<%
	for absrecordnum = 1 to recordperpage
%>

		<tr onmouseover="bgColor='#EEEEEE';" onmouseout="bgColor='';">
			<td class=ctd><%=icounter%></td>
			<td class=ltd><span alt="<%=replace(replace(replace(Rs("kp_bz"),vbcrlf,"<br>"),"'","&#39;"),Chr(34),"&#34;") & "<br><br><b>������</b>:" & Rs("kp_kpr")%>"><%=Rs("kp_item")%>&nbsp;</span></td>
			<td class=ctd><%=Formatnumber(Rs("kp_uprice"),1,-1)%></td>
			<td class=ctd><%=Rs("kp_mul")%>&nbsp;</td>
			<td class=ctd><%=Rs("kp_zrr")%></td>
			<td class=ctd><%=Rs("kp_zrrjs")%>&nbsp;</td>
			<td class=ctd><%=xjDate(rs("kp_time"),1)%></td>
			<%If ChkAble("1,2,3,11,12") Then%>
			<td class=ctd>
				&nbsp;
				<%If Rs("kp_kpr")=Session("userName") or ChkAble("3") Then%><a href="ygkp_change.asp?id=<%=Rs("ID")%>&kind=<%=Rs("kp_kind")%>&group=<%=Rs("kp_group")%>&ipage=<%=striPage&strFeedBack%>">����</a><%End If%>
				<%If ChkAble(3) or (Rs("kp_kpr")=Session("userName") and Session("userdepart")="Ʒ�ܲ�") Then%>
				 | <a href="ygkp_delete.asp?id=<%=Rs("ID")%>&ipage=<%=striPage&strFeedBack%>" onclick="return confirm('ɾ���󽫲��ָܻ�!\n\nȷ��ɾ�� �� <%=iCounter%> �� ������Ϣ��?');">ɾ��</a><%End If%>
				&nbsp;
			</td>
			<%End If%>
		</tr>
		<%rs.movenext%>
		<%if rs.eof then%>
			<%exit for%>
		<%end if%>
		<%icounter = icounter - 1%>
	<%next%>
	</table>
	<table width="95%" cellpadding=2 cellspacing=0 border=0  align="center">
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