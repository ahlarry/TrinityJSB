<!--#include file="include/conn.asp"-->
<!--#include file="include/page/ftask_dbinf.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="�������� �� ���������б�"
strPage="ftask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

Dim strFeedBack,strTerm,strType,strxldh,strxlxh,strxlcj,ifz
strTerm=Trim(Request("term"))
strType=Trim(Request("type"))
strxldh=Trim(Request("xldh"))
strxlxh=Trim(Request("xlxh"))
strxlcj=Trim(Request("xlcj"))
ifz=(Request("fz"))
strFeedBack="&term="&strTerm&"&type="&strType&"&xldh="&strxldh&"&xlxh="&strXlxh&"&xlcj="&strXlcj&"&fz="&ifz

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchInfo()%>
    </td>
  </tr>
  <Tr>
    <Td class=ctd height=300><%Call ftaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</Table>
<%
End Sub
Function Task_Ts(mystr)
       dim strnew
		if instr(mystr,"||")>0then
		strnew=split(mystr,"||")
	    for i=0 to ubound(strnew)
		Task_Ts=Task_Ts&strnew(i)&"<br>"
		next
		else
		Task_Ts=mystr
		end if

End Function


Function SearchInfo()
%>
<table>
  <tr>
    <td></td>
    <td></td>
  </tr>
</table>
<Table border=0 cellpadding=2 cellspacing=0 width="100%">
  <Form action="<%=Request.Servervariables("script_name")%>" method=get>
    <Tr>
      <Td>&nbsp;&nbsp;<strong>����&nbsp;����</strong>&nbsp;������:</td>
      <Td><input type="text" name="xldh" size="12" value=<%=strxldh%>></td>
      <Td>����С��:
        <input type="text" name="xlxh" size="10" value=<%=strxlxh%>>
      </td>
      <Td> ������:
        <input type="text" name="xlcj" size="10" value=<%=strxlcj%>></td>
      <Td> ��ֵ:
        <input type="text" name="fz" size="5" value=<%=ifz%>></td>
      <td rowspan="2"><input type=submit value=" ѡ�� ">
      </Td>
    </Tr>
    <tr>
      <td>&nbsp;&nbsp;<strong>ɸѡ&nbsp;����</strong>&nbsp;&nbsp;������:</td>
      <Td><Select name="term" onchange='window.location.href("<%=request.servervariables("script_name")%>?type=<%=strType%>&xldh=<%=strxldh%>&xlxh=<%=strxlxh%>&xlcj=<%=strxlcj%>&fz=<%=ifz%>&term=" + this.value);'>
          <option value="">ȫ��</option>
          <%If Session("userName")<>"" Then %>
          <option value="<%=Session("userName")%>"><%=Session("userName")%></option>
          <%End If%>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>' <%If strTerm=c_jsb(i) Then%> Selected<%End If%>><%=c_jsb(i)%></option>
          <%next%>
        </Select>
      </td>
      <Td> ��������:
        <select name="type" onchange='location.href("<%=request.servervariables("script_name")%>?term=<%=strterm%>&xldh=<%=strxldh%>&xlxh=<%=strxlxh%>&xlcj=<%=strxlcj%>&fz=<%=ifz%>&type=" + this.value);'>
          <option value="">ȫ��</option>
          <%for i = 0 to ubound(c_lxrwlx)%>
          <option value='<%=c_lxrwlx(i)%>'<%If strType=c_lxrwlx(i) Then%> Selected<%End If%>><%=c_lxrwlx(i)%></option>
          <%next%>
        </select>
      </td>
      <td colspan="2">&nbsp;&nbsp;</td>
    </tr>
  </Form>
</Table>
<%
End Function

Function ftaskList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter, TotalCount, strTerm, strType, strSql,strxldh,strXlxh,strXlcj,ifz
	absPageNum = 0
	RecordPerPage = 20
	strSql = "" : strxldh="":ifz=""
	strTerm=Trim(Request("term"))
	strType=Trim(Request("type"))
	strxldh=Trim(Request("xldh"))
	strXlxh=Trim(Request("xlxh"))
	strXlcj=Trim(Request("Xlcj"))
	ifz=(Request("fz"))

	If strTerm <> "" Then
		strSql = " zrr='"&strTerm&"' "
	End If

	If strType <> "" Then
		If strSql <> "" Then
			strSql = "rwlx like '%"&strType&"%' and " & strSql
		Else
			strSql  = "rwlx like '%"&strType&"%' "
		End If
	End If

	If strxldh <> "" Then
		If strSql <> "" Then
			strSql = "xldh='"&strxldh&"' and " & strSql
		Else
			strSql  = "xldh='"&strxldh&"' "
		End If
	End If

	If strXlxh <> "" Then
		If strSql <> "" Then
			strSql = "rwlr like '%"&strxlxh&"%' and " & strSql
		Else
			strSql  = "rwlr like '%"&strxlxh&"%' "
		End If
	End If

	If strXlcj <> "" Then
		If strSql <> "" Then
			strSql = "rwlr like '%"&strXlcj&"%' and " & strSql
		Else
			strSql  = "rwlr like '%"&strXlcj&"%' "
		End If
	End If

	If ifz<>"" Then
		If strSql <> "" Then
			strSql = "zf="&ifz&" and " & strSql
		Else
			strSql  = "zf="&ifz&""
	    End If
	End If

	If strSql <> "" Then
		strSql = "select * from [ftask] where "&strSql & " order by jssj desc"
	Else
		strSql = "select * from [ftask] order by jssj desc"
	End If
	Set Rs = Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Call xjweb.Exec("",-1)
	Rs.open strSql,Conn,3,3
	If Rs.Eof Or Rs.Bof Then
		Call TbTopic(strTerm & " û���κ�"&strType&"����") : Exit Function
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
	Call TbTopic("����ģ�߳����������б�")

	'icounter = (absPageNum - 1) * recordperpage + 1
%>
<table width="95%" cellpadding=2 cellspacing=0 class=xtable align="center">
  <tr>
    <th class=th width=40>ID</th>
    <th class=th width=*>��������</th>
    <th class=th width=*>�������</th>
    <th class=th width=100>��ֵ</th>
    <th class=th width=100>������</th>
    <%if chkable(3) then%>
    <th class=th width=120 colspan=2>��  ��</th>
    <%end if%>
  </tr>
  <%
	for absrecordnum = 1 to recordperpage
	call Task_Ts(rs("rwlr"))
%>
  <tr>
    <td class=ctd><%=icounter%></td>
    <td class=ctd <%if Rs("xldh")<>"" then%>alt="������:<%=rs("xldh")%><br><%=Task_Ts(rs("rwlr"))%>"<%else%>alt="��������:<br><%=Task_Ts(rs("rwlr"))%>"<%end if%>><%=rs("rwlx")%></td>
    <td class=ctd><%=xjDate(rs("jssj"),1)%></td>
    <td class=ctd><%=rs("zf")%></td>
    <td class=ctd><%=rs("zrr")%></td>
    <%if chkable(3) then%>
    <td class=ctd width=60><a href="ftask_change.asp?id=<%=rs("id")%>" target="_blank">����</a></td>
    <form action="ftask_indb.asp?action=delete" method=post onsubmit="return confirm('ɾ��ID��Ϊ <%=icounter%> ���������� ɾ���󽫲��ָܻ���\nȷ����');">
      <td class=ctd width=60><input type="submit" value="ɾ��"></td>
      <input type="hidden" name=id value="<%=rs("id")%>">
    </form>
    <%end if%>
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
    <td align=left> ���������� <%=rs.recordcount%> ��&nbsp;&nbsp;
      ÿҳ <%=rs.pagesize%> ��&nbsp;&nbsp;
      �� <%=Rs.PageCount%> ҳ&nbsp;&nbsp;
      ��ǰΪ�� <%=absPageNum%> ҳ </td>
    <td align=right> ��
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
