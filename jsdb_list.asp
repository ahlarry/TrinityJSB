<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'14:32 2007-1-22-����һ
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="������� �� ��������"
strPage="mtask"
xjweb.header()
Call TopTable()

Dim strFeedBack, strHth, strKhmc, strzuz, strOrder, strTerm, striPage
strHth=Trim(Request("hth"))
strKhmc=Trim(Request("khmc"))
strzuz=Trim(Request("zuz"))
strOrder=Trim(Request("order"))
strTerm=Trim(Request("term"))
striPage =request("ipage")

	strFeedBack=""
	If strHth<>"" Then strFeedBack="&hth="&strHth
	If strKhmc<>"" Then strFeedBack="&khmc="&strKhmc&strFeedBack
	If strzuz<>"" Then strFeedBack="&zuz="&strzuz&strFeedBack
	If strOrder<>"" Then strFeedBack="&order="&strOrder&strFeedBack
	If strTerm<>"" Then strFeedBack="&term="&strTerm&strFeedBack

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()

%>
<table class="xtable" cellspacing="0" cellpadding="2" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd"><%Call SearchInfo()%></td>
  </tr>
  <tr>
    <td class="ctd" height="280"><%Call TaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%></td>
  </tr>
</table>
<%
End Sub

Function SearchInfo()
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td>&nbsp;&nbsp;��ͬ��:
        <input type="text" name="hth" value="<%=strHth%>" size="15" /></td>
      <td>�ͻ�����:
        <input type="text" name="khmc" value="<%=strKhmc%>" size="8" /></td>
      <td>�鳤:
        <input type="text" name="zuz" value="<%=strzuz%>" size="8" /></td>
      <td align="center"><input type="submit" value=" �� �� �� �� " /></td>
      <td>����:
        <select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;hth=<%= strhth%>&amp;khmc=<%=strKhmc%>&amp;term=<%=strterm%>&amp;order=&quot; + this.value);'>
          <%If strOrder = "hth" Then%>
          <option value="jhjssj">�ƻ����ʱ��</option>
          <option value="hth" selected="selected">��ͬ��</option>
          <option value="zz">�鳤</option>
          <%ElseIf strOrder="zz" Then%>
          <option value="jhjssj">�ƻ����ʱ��</option>
          <option value="hth">��ͬ��</option>
          <option value="zz" selected="selected">�鳤</option>
          <%Else%>
          <option value="jhjssj" selected="selected">�ƻ����ʱ��</option>
          <option value="hth">��ͬ��</option>
          <option value="zz">�鳤</option>
          <%End If%>
        </select></td>
      <td>����:
        <select name="term" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;hth=<%= strhth%>&amp;khmc=<%=strkhmc%>&amp;order=<%=strorder%>&amp;term=&quot; + this.value);'>
          <%select case strterm%>
          <%case "no"%>
          <option value="no" selected="selected">δ���</option>
          <option value="ok">�����</option>
          <option value="all">ȫ��</option>
          <%case "ok"%>
          <option value="no">δ���</option>
          <option value="ok" selected="selected">�����</option>
          <option value="all">ȫ��</option>
          <%case else%>
          <option value="no">δ���</option>
          <option value="ok">�����</option>
          <option value="all" selected="selected">ȫ��</option>
          <%end select%>
        </select></td>
    </tr>
  </form>
</table>
<%
End Function

Function TaskList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter
	absPageNum = 0
	strSql = ""
	RecordPerPage=20 		'ÿҳ��ʾ��¼��
	If strHth <> "" Then
			strSql  = " hth like '%"&strHth&"%' "
	End If
	If strkhmc <> "" Then
		If strSql <> "" Then
			strSql = " khmc like '%"&strkhmc&"%' and " & strSql
		Else
			strSql  = " khmc like '%"&strkhmc&"%' "
		End If
	End If
	If strzuz <> "" Then
		If strSql <> "" Then
			strSql = " zz like '%"&strzuz&"%'  and " & strSql
		Else
			strSql  = " zz like '%"&strzuz&"%' "
		End If
	End If

	Select Case strTerm
			Case "no"
				If strSql <> "" Then
					strSql = " isnull(shjssj) and " & strSql
				Else
					strSql  = " isnull(shjssj) "
				End If
			Case "ok"
				If strSql <> "" Then
					strSql = " not(isnull(shjssj)) and " & strSql
				Else
					strSql  = " not(isnull(shjssj)) "
				End If
			Case Else
	End Select

	If Request("strSql") <> "" Then
		strSql = request("strSql")
	End If

	Dim sqlOrder
	sqlOrder = " order by jhjssj desc, hth desc"
	If LCase(strOrder) = "jhjssj" Then sqlOrder = " order by jhjssj desc, hth desc"
	If LCase(strOrder) = "hth" Then sqlOrder = " order by hth desc,zz desc"
	If LCase(strOrder) = "zz" Then sqlOrder = " order by zz desc, hth desc, jhjssj desc"
	If strSql <> "" Then
		strSql = "select * from [jsdb] where "& strSql & sqlOrder
	Else
		strSql = "select * from [jsdb] "& sqlOrder
	End If
	Set Rs = Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Call xjweb.Exec("",-1)
	Rs.open strSql,Conn,3,3
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("ָ�������������鲻����,��ſ�ɸѡ����!","")
	End If
	Rs.PageSize = RecordPerPage

	If Trim(Request("iPage")) <> ""  Then
		If IsNumeric(Trim(Request("iPage"))) Then
			If Trim(Request("iPage")) <= 0 Then
				absPageNum=1
			ElseIf CLng(Trim(Request("iPage"))) > Rs.Pagecount Then
				absPageNum = Rs.Pagecount
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

	If absPageNum>Rs.PageCount Then absPageNum=Rs.PageCount
	Rs.absolutePage = absPageNum
	Call CutLine()		'��ʾͼ��
	Call TbTopic("����ģ�߳��������������б�")
	iCounter = (absPageNum - 1) * RecordPerPage + 1
%>
<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable" align="center">
  <tr>
    <th class="th" width="25">id </th>
    <th class="th" width="60">��ͬ�� </th>
    <th class="th" width="80">�ͻ����� </th>
    <th class="th" width="*">�������� </th>
    <th class="th" width="40">�鳤 </th>
    <th class="th" width="80">�ƻ���� </th>
    <th class="th" width="*">��� </th>
    <th class="th" width="*">��� </th>
  </tr>
  <%
  	For absRecordNum = 1 To RecordPerPage		'��ҳ
%>
  <tr>
    <td class="ctd"><%=iCounter%></td>
    <td class="ctd"><%=rs("hth")%></a></td>
    <td class="ctd" ><%=rs("khmc")%></a></td>
    <td class="ctd" alt="<%=rs("rwnr")%>"><%=xjweb.StringCut(rs("rwnr"),16)%></td>
    <td class="ctd"><%=rs("zz")%></td>
    <td class="ctd"><%=rs("jhjssj")%></td>
    <td class="ctd"><%call distd(rs("sjkssj"),rs("sjjssj"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("shkssj"),rs("shjssj"),0,rs)%></td>
    <%rs.movenext%>
    <%if rs.eof then
  		exit for
  	end if
  iCounter=iCounter+1
  next%>
</table>
<table width="98%" cellpadding="2" cellspacing="0" border="0" align="center">
  <tr>
    <td align="left"> ���������� <%=Rs.RecordCount%> ��&nbsp;&nbsp;
      ÿҳ <%=Rs.PageSize%> ��&nbsp;&nbsp;
      �� <%=Rs.PageCount%> ҳ&nbsp;&nbsp;
      ��ǰΪ�� <%=absPageNum%> ҳ </td>
    <td align="right"> ��
      <%
				If absPageNum>1 Then
					Response.Write("<a href="&request.servervariables("script_name")&"?ipage="&(abspagenum-1)&strFeedBack&" title='��һҳ'> ��</a>&nbsp;&nbsp;")
				End If
				Dim iStart,iEnd
				If absPageNum<4 Then
					iStart=1
				Else
					iStart=absPageNum-3
				End If
				If absPageNum<Rs.PageCount-3 Then
					iEnd = absPageNum + 3
				Else
					iEnd = Rs.PageCount
				End If
				For i=iStart to iEnd
					If i=absPageNum Then
						Response.Write("&nbsp;<font style=font-size:11pt;><b>"&  i & "</b></font>&nbsp;")
					Else
						Response.Write("&nbsp;<a href="&Request.ServerVariables("SCRIPT_NAME")&"?iPage="&i&strFeedBack&">" & i & "</a>&nbsp;")
					End If
				Next
				If absPageNum<Rs.PageCount Then
					Response.Write("&nbsp;<a href="&Request.ServerVariables("SCRIPT_NAME")&"?iPage="&(abspagenum+1)&strFeedBack&" title='��һҳ'> �� </a>&nbsp;")
				End If
			%>
      ��
      ��ת��:
      <select name="ipage" onchange='location.href(&quot;<%=Request.ServerVariables("SCRIPT_NAME")%>?ipage=&quot; + this.value +&quot;<%=strFeedback%>&quot;);'>
        <%For i=1 To Rs.PageCount%>
        <%If i=absPageNum Then%>
        <option value="<%=i%>" selected="selected">�� <%=i%> ҳ</option>
        <%Else%>
        <option value="<%=i%>">�� <%=i%> ҳ</option>
        <%End If%>
        <%Next%>
      </select></td>
  </tr>
</table>
<%
End Function
%>
