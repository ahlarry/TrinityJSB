<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="ģ�ߵ��� �� ������Ϣ�б�"
strPage="mtest"
'Call FileInc(0, "js/ftask.js")
xjweb.header()
Call TopTable()

Dim strFeedBack, strlsh, strddh, strdwmc, strdmmc, strorder, strterm, strtsy, strsjs
strlsh = Trim(Request("lsh"))
strddh = Trim(Request("ddh"))
strdwmc = Trim(Request("dwmc"))
strdmmc = Trim(Request("dmmc"))
strorder = Trim(Request("order"))
If strorder="" Then strorder="tsgxsj"
strterm = Trim(Request("term"))
strsjs = Trim(Request("sjs"))
strtsy = Trim(Request("tsy"))
strFeedBack = "&lsh="&strlsh&"&ddh="&strddh&"&dwmc="&strdwmc&"&dmmc="&strdmmc&"&order="&strorder&"&term="&strterm&"&sjs="&strsjs&"&tsy="&strtsy

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchInfo()%></Td>
  </Tr>
  <Tr>
    <Td class=ctd height=300><%Call mtestList()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function SearchInfo()
%>
<table width="100%" cellpadding=2 cellspacing=0 border=0 height="100%">
  <form action="<%=Request.servervariables("script_name")%>" method="get">
    <tr>
      <td height="25"> ��ˮ��:
        <input type="text" name="lsh" value="<%=strlsh%>" size="6">
        ������:
        <input type="text" name="ddh" value="<%=strddh%>" size="12">
        ��λ����:
        <input type="text" name="dwmc" value="<%=strdwmc%>" size="8">
        ��������:
        <input type="text" name="dmmc" value="<%=strdmmc%>" size="8">
        <input type="submit" value=" ���� "><br>
        ���Խ��:
        <select name="term" onchange='location.href("<%=Request.servervariables("script_name")%>?ipage=1&lsh=<%= strlsh%>&ddh=<%=strddh%>&dwmc=<%=strdwmc%>&order=<%=strorder%>&sjs=<%=strsjs%>&tsy=<%=strtsy%>&term=" + this.value);'>
          <option value="all" selected>ȫ��</option>
          <option value="no" <%If strterm="no" Then%>selected<%End If%>>���ڵ���</option>
          <option value="ok" <%If strterm="ok" Then%>selected<%End If%>>�������</option>
          <option value="hg" <%If strterm="hg" Then%>selected<%End If%>>���Ժϸ�</option>
          <option value="ct" <%If strterm="ct" Then%>selected<%End If%>>���ڳ���</option>
          <option value="jt" <%If strterm="jt" Then%>selected<%End If%>>���⾫��</option>
          <option value="jy" <%If strterm="jy" Then%>selected<%End If%>>��������</option>
          <option value="lc" <%If strterm="lc" Then%>selected<%End If%>>��������</option>
        </select>&nbsp;&nbsp;&nbsp;&nbsp;
		�ṹ���ʦ:
        <select name="sjs" onchange='location.href("<%=Request.servervariables("script_name")%>?ipage=1&lsh=<%= strlsh%>&ddh=<%=strddh%>&dwmc=<%=strdwmc%>&order=<%=strorder%>&term=<%=strterm%>&tsy=<%=strtsy%>&sjs=" + this.value);'>
          <option value="" selected>ȫ��</option>
          <%for i = 0 to ubound(c_jsb)%>
          	<option value="<%=c_jsb(i)%>"<%If strsjs=c_jsb(i) Then%> Selected<%End If%>><%=c_jsb(i)%></option>
		  <%next%>
        </select>&nbsp;&nbsp;&nbsp;&nbsp;
		���Թ���ʦ:
        <select name="tsy" onchange='location.href("<%=Request.servervariables("script_name")%>?ipage=1&lsh=<%= strlsh%>&ddh=<%=strddh%>&dwmc=<%=strdwmc%>&order=<%=strorder%>&term=<%=strterm%>&sjs=<%=strsjs%>&tsy=" + this.value);'>
          <option value="" selected>ȫ��</option>
          <%for i = 0 to ubound(c_xz5)%>
          	<option value="<%=c_xz5(i)%>"<%If strtsy=c_xz5(i) Then%> Selected<%End If%>><%=c_xz5(i)%></option>
		  <%next%>
        </select>&nbsp;&nbsp;&nbsp;&nbsp;
        ����:
        <select name="order" onchange='location.href("<%=Request.servervariables("script_name")%>?ipage=1&lsh=<%= strlsh%>&ddh=<%=strddh%>&dwmc=<%=strdwmc%>&term=<%=strterm%>&sjs=<%=strsjs%>&tsy=<%=strtsy%>&order=" + this.value);'>
          <option value="lsh" selected="selected">��ˮ��</option>
          <option value="ddh" <%If strOrder="ddh" Then%>selected<%End If%>>������</option>
          <option value="tscs" <%If strOrder="tscs" Then%>selected<%End If%>>���Դ���</option>
          <option value="tskssj" <%If strOrder="tskssj" Then%>selected<%End If%>>��ʼ����</option>
          <option value="tsgxsj" <%If strOrder="tsgxsj" Then%>selected<%End If%>>��������</option>
          <option value="tsjssj" <%If strOrder="tsjssj" Then%>selected<%End If%>>�������</option>
        </select></td>
    </tr>
  </form>
</table>
<%
End Function

function mtestList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCount
	absPageNum = 0
	strSql = ""
	RecordPerPage = 20
	if strlsh <> "" then
		strSql = " a.lsh = '"&strlsh&"' "
	end if

	if strdwmc <> "" then
		if strSql <> "" then
			strSql = " dwmc like '%"&strdwmc&"%' and " & strSql
		else
			strSql  = " dwmc like '%"&strdwmc&"%' "
		end if
	end if

	if strdmmc <> "" then
		if strSql <> "" then
			strSql = " dmmc like '%"&strdmmc&"%' and " & strSql
		else
			strSql  = " dmmc like '%"&strdmmc&"%' "
		end if
	end if

	if strddh <> "" then
		if strSql <> "" then
			strSql = " ddh like '%"&strddh&"%' and " & strSql
		else
			strSql  = " ddh like '%"&strddh&"%' "
		end if
	end if

	if strsjs <> "" then
		if strSql <> "" then
			strSql = " (mtjgr = '"&strsjs&"' or dxjgr = '"&strsjs&"' or gjjgr = '"&strsjs&"')  and " & strSql
		else
			strSql = " (mtjgr = '"&strsjs&"' or dxjgr = '"&strsjs&"' or gjjgr = '"&strsjs&"') "
		end if
	end if

	if strtsy <> "" then
		if strSql <> "" then
			strSql = " (mttsr = '"&strtsy&"' or dxtsr = '"&strtsy&"')  and " & strSql
		else
			strSql  = " (mttsr = '"&strtsy&"' or dxtsr = '"&strtsy&"') "
		end if
	end if

	select case strterm
			case "no"
				if strSql <> "" then
					strSql = " isnull(tsjssj) and " & strSql
				else
					strSql  = " isnull(tsjssj) "
				end if
			case "ok"
				if strSql <> "" then
					strSql = " not(isnull(tsjssj)) and " & strSql
				else
					strSql  = " not(isnull(tsjssj)) "
				end if
			case "hg"
				if strSql <> "" then
					strSql = " tsjg='���Ժϸ�' and " & strSql
				else
					strSql  = " tsjg='���Ժϸ�' "
				end if
			case "ct"
				if strSql <> "" then
					strSql = " tsjg='���ڳ���' and " & strSql
				else
					strSql  = " tsjg='���ڳ���' "
				end if
			case "jt"
				if strSql <> "" then
					strSql = " tsjg='���⾫��' and " & strSql
				else
					strSql  = " tsjg='���⾫��' "
				end if
			case "jy"
				if strSql <> "" then
					strSql = " tsjg='Ԥ���ջ����' and " & strSql
				else
					strSql  = " tsjg='Ԥ���ջ����' "
				end if
			case "lc"
				if strSql <> "" then
					strSql = " tsjg='��������' and " & strSql
				else
					strSql  = " tsjg='��������' "
				end if
			case else
	end select

	if Request("strSql") <> "" then
		strSql = Request("strSql")
	end if

	Dim sqlorder
	if lcase(strorder) = "lsh" then sqlorder = " order by a.lsh desc"
	if lcase(strorder) = "ddh" then sqlorder = " order by ddh desc, a.lsh desc"
	if lcase(strorder) = "tscs" then sqlorder = " order by b.tscs desc"
	if lcase(strorder) = "tskssj" then sqlorder = " order by b.tskssj desc"
	if lcase(strorder) = "tsgxsj" then sqlorder = " order by b.tsgxsj desc"
	if lcase(strorder) = "tsjssj" then sqlorder = " order by b.tsjssj desc"

	if strSql <> "" then
		strSql = "select a.lsh, a.ddh, a.dwmc, a.dmmc, a.bz, b.*, a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh=b.lsh and " &strSql & sqlorder
	else
		strSql = "select a.lsh, a.ddh, a.dwmc, a.dmmc, a.bz, b.*, a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh=b.lsh " & sqlorder
	end if
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	rs.cachesize = RecordPerPage
	rs.open strSql,conn,3,3
	if rs.eof or rs.bof then
		Call JsAlert("ָ��������û���κ������飬������ѡ��������","mtest_list.asp")
		Exit Function
	end if
	rs.pagesize = RecordPerPage

	if Trim(Request("iPage")) <> ""  then
		if IsNumeric(Trim(Request("iPage"))) then
			if Trim(Request("iPage")) <= 0 then
				absPageNum = 1
			elseif CLng(Trim(Request("iPage"))) > rs.pagecount then
				absPageNum = rs.pagecount
			else
				absPageNum = CLng(Trim(Request("iPage")))
			end if
		else
			if Request("iCurPage") <> "" then
				absPageNum = CLng(Request("iCurPage"))
			else
				absPageNum = 1
			end if
		end if
	else
		if Request("iCurPage") <> "" then
			absPageNum = CLng(Request("iCurPage"))
		else
			absPageNum = 1
		end if
	end if

	if absPageNum > rs.pagecount then absPageNum = rs.pagecount
	rs.absolutepage = absPageNum

	Call TbTopic("����ģ�߳�������Ϣ�б�")
	iCount = (absPageNum - 1) * RecordPerPage + 1
%>
<table width="98%" cellpadding=2 cellspacing=0 class=xtable align="center">
  <tr>
    <th class=th width=25>id</th>
    <th class=th width=60>������</th>
    <th class=th width=50>��ˮ��</th>
    <th class=th width=80>��λ����</th>
    <th class=th width=*>��������</th>
    <th class=th width=100>��ʼ����</th>
    <th class=th width=100>��������</th>
    <th class=th width=100>�������</th>
    <th width=80 colspan="2" class=th>���Դ���</th>
  </tr>
  <%
	Dim ilsh, TmpSql, Tmprs, itslb, iedsx, iedxx
	for absRecordNum = 1 to RecordPerPage
		ilsh=rs("lsh")
		TmpSql="select * from [mtask] where lsh='"&ilsh&"'"
		set Tmprs=xjweb.Exec(TmpSql, 1)
		itslb=Tmprs("tslb")
		Tmprs.Close

		TmpSql="select * from [c_tscs] where dmlb='"&itslb&"'"
		set Tmprs=xjweb.Exec(TmpSql, 1)
			If not(Tmprs.Eof Or Tmprs.Bof) Then
				iedsx=Tmprs("edsx")
				iedxx=Tmprs("edxx")
			else
				iedsx=0
				iedxx=0
			End If
		Tmprs.Close
%>
  <tr>
    <td class=ctd><%=iCount%></td>
    <td class=ctd><%=rs("ddh")%></td>
    <td class=ctd><a href=mtest_display.asp?s_lsh=<%=rs("lsh")%>>
    		<%If InStr(rs("bz"),"���Թ�ע")>0 Then
       		Response.Write("<font color=red><b>"&rs("lsh")&"</b></font>")
    		Else
      		Response.Write(rs("lsh"))
     	End If%>
    </a></td>
    <td class=ctd><%=rs("dwmc")%></td>
    <td class=ctd><%=rs("dmmc")%></td>
    <td class=ctd><%=xjDate(rs("tskssj"),1)%></td>
    <td class=ctd><%=xjDate(rs("tsgxsj"),1)%></td>
    <td class=ctd><%=xjDate(rs("tsjssj"),1)%></td>
    <td class=ctd><%If rs("tscs")>iedsx and iedxx<>0 Then
				Response.Write("<font color='#ff0000'><strong>"&rs("tscs")&"</strong></font>")
			else if rs("tscs")<iedxx and not(isnull(rs("tsjssj"))) Then
					Response.Write("<font color='#8000ff'><strong>"&rs("tscs")&"</strong></font>")
				else
					Response.Write(rs("tscs"))
				End If
			End If%></td>
    <td class=ctd><%If rs("tscs")>iedsx and iedxx<>0 Then
				Response.Write("<font color='#ff0000'><strong>"&iedxx&"-"&iedsx&"</strong></font>")
			else if rs("tscs")<iedxx and not(isnull(rs("tsjssj"))) Then
					Response.Write("<font color='#8000ff'><strong>"&iedxx&"-"&iedsx&"</strong></font>")
				else
					Response.Write(iedxx&"-"&iedsx)
				End If
			End If%></td>
  </tr>
  <%rs.movenext%>
  <%if rs.eof then%>
  <%exit for%>
  <%end if%>
  <%iCount = iCount + 1%>
  <%next%>
</table>
<table width="95%" cellpadding=2 cellspacing=0 border=0 align="center">
  <tr>
    <td align=left> ���������� <%=rs.recordcount%> ��&nbsp;&nbsp;
      ÿҳ <%=rs.pagesize%> ��&nbsp;&nbsp;
      �� <%=rs.pagecount%> ҳ&nbsp;&nbsp;
      ��ǰΪ�� <%=absPageNum%> ҳ </td>
    <td align=right> ��
      <%
				if absPageNum > 1 then
					Response.write("<a href="&Request.servervariables("script_name")&"?ipage="&(abspagenum-1)&strFeedBack&" title='��һҳ'> ��</a>&nbsp;&nbsp;")
				end if
				dim iStart,iEnd
				if absPageNum < 4 then
					iStart = 1
				else
					iStart = absPageNum - 3
				end if
				if absPageNum < rs.pagecount - 3 then
					iEnd = absPageNum + 3
				else
					iEnd = rs.pagecount
				end if
				for i = iStart to iEnd
					if i = absPageNum then
						Response.write("&nbsp;<font style=font-size:11pt;><b>"&  i & "</b></font>&nbsp;")
					else
						Response.write("&nbsp;<a href="&Request.servervariables("script_name")&"?ipage="&i&strFeedBack&">" & i & "</a>&nbsp;")
					end if
				next
				if absPageNum < rs.pagecount then
					Response.write("&nbsp;<a href="&Request.servervariables("script_name")&"?ipage="&(absPageNum+1)&strFeedBack&" title='��һҳ'> �� </a>&nbsp;")
				end if
			%>
      ��
      ��ת��:
      <select name="ipage" onchange='location.href("<%=Request.servervariables("script_name")%>?ipage=" + this.value +"<%=strFeedBack%>");'>
        <%for i=1 to rs.pagecount%>
        <%if i = absPageNum then%>
        <option value=<%=i%> selected>�� <%=i%> ҳ</option>
        <%else%>
        <option value=<%=i%>>�� <%=i%> ҳ</option>
        <%end if%>
        <%next%>
      </select></td>
  </tr>
</table>
<%
end function
%>
