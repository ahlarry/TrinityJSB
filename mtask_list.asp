<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'14:32 2007-1-22-����һ
Call ChkPageAble(0)
'Call ChkDepart("������")
CurPage="������� �� ��������"
strPage="mtask"
xjweb.header()
Call TopTable()

Dim action, strFeedBack, strLsh, strDdh, strDwmc, strDmmc, strseb, strjcj, strqshu, strzuz, strdxjs, strOrder, strTerm, striPage, strkhlx
action=Request("action")
strLsh=Trim(Request("lsh"))
strDdh=Trim(Request("ddh"))
strDwmc=Trim(Request("dwmc"))
strDmmc=Trim(Request("dmmc"))
strseb=Trim(Request("seb"))
strjcj=Trim(Request("jcj"))
strqshu=Trim(Request("qshu"))
strdxjs=Trim(Request("dxjs"))
strzuz=Trim(Request("zuz"))
strOrder=Trim(Request("order"))
strTerm=Trim(Request("term"))
striPage =request("ipage")

	strFeedBack=""
	If strLsh<>"" Then strFeedBack="&lsh="&strLsh
	If strDdh<>"" Then strFeedBack="&ddh="&strDdh&strFeedBack
	If strDwmc<>"" Then strFeedBack="&dwmc="&strDwmc&strFeedBack
	If strDmmc<>"" Then strFeedBack="&dmmc="&strDmmc&strFeedBack
	If strseb<>"" Then strFeedBack="&seb="&strseb&strFeedBack
	If strjcj<>"" Then strFeedBack="&jcj="&strjcj&strFeedBack
	If strqshu<>"" Then strFeedBack="&qshu="&strqshu&strFeedBack
	If strzuz<>"" Then strFeedBack="&zuz="&strzuz&strFeedBack
	If strdxjs<>"" Then strFeedBack="&dxjs="&strdxjs&strFeedBack
	If strOrder<>"" Then strFeedBack="&order="&strOrder&strFeedBack
	If strTerm<>"" Then strFeedBack="&term="&strTerm&strFeedBack

	If action="KhChan" and ChkAble(4) and Session("userGroup")=5 Then
		Call KhChan()
	End If

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
      <td>&nbsp;&nbsp;������:
        <input type="text" name="ddh" value="<%=strDdh%>" size="10" /></td>
      <td>��ˮ��:
        <input tabindex="1" type="text" name="lsh" size="8" value="<%=strLsh%>" /></td>
      <td>��λ:
        <input type="text" name="dwmc" value="<%=strDwmc%>" size="8" /></td>
      <td>����:
        <input type="text" name="dmmc" value="<%=strDmmc%>" size="8" /></td>
      <td>����:
        <select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;lsh=<%= strlsh%>&amp;ddh=<%=strddh%>&amp;dwmc=<%=strdwmc%>&amp;term=<%=strterm%>&amp;seb=<%=strseb%>&amp;jcj=<%=strjcj%>&amp;qshu=<%=strqshu%>&amp;order=&quot; + this.value);'>
          <%If strOrder = "lsh" Then%>
          <option value="jhjssj">�ƻ����ʱ��</option>
          <option value="ddh">������</option>
          <option value="lsh" selected="selected">��ˮ��</option>
          <option value="zz">�鳤</option>
          <%ElseIf strOrder="ddh" Then%>
          <option value="jhjssj">�ƻ����ʱ��</option>
          <option value="ddh" selected="selected">������</option>
          <option value="lsh">��ˮ��</option>
          <option value="zz">�鳤</option>
          <%ElseIf strOrder="zz" Then%>
          <option value="jhjssj">�ƻ����ʱ��</option>
          <option value="ddh">������</option>
          <option value="lsh">��ˮ��</option>
          <option value="zz" selected="selected">�鳤</option>
          <%Else%>
          <option value="jhjssj" selected="selected">�ƻ����ʱ��</option>
          <option value="ddh">������</option>
          <option value="lsh">��ˮ��</option>
          <option value="zz">�鳤</option>
          <%End If%>
        </select></td>
      <td>����:
        <select name="term" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;lsh=<%= strlsh%>&amp;ddh=<%=strddh%>&amp;dwmc=<%=strdwmc%>&amp;seb=<%=strseb%>&amp;jcj=<%=strjcj%>&amp;qshu=<%=strqshu%>&amp;order=<%=strorder%>&amp;term=&quot; + this.value);'>
          <%select case strterm%>
          <%case "no"%>
          <option value="no" selected="selected">δ���</option>
          <option value="ok">�����</option>
          <option value="all">ȫ��</option>
          <%case "ok"%>
          <option value="no">δ���</option>
          <option value="ok" selected="selected">�����</option>
          <option value="all">ȫ��</option>
          <%case "all"%>
          <option value="no">δ���</option>
          <option value="ok">�����</option>
          <option value="all" selected="selected">ȫ��</option>
          <%case else%>
          <option value="no">δ���</option>
          <option value="ok">�����</option>
          <option value="all" selected="selected">ȫ��</option>
          <%end select%>
        </select></td>
    </tr>
    <tr>
      <td>&nbsp;&nbsp;�豸��:
        <input type="text" name="seb" value="<%=strseb%>" size="10" /></td>
      <td>������:
        <input type="text" name="jcj" value="<%=strjcj%>" size="8" /></td>
      <td>ǻ��:
        <input type="text" name="qshu" value="<%=strqshu%>" size="8" /></td>
      <td>��ע:
        <input type="text" name="dxjs" value="<%=strdxjs%>" size="8" /></td>
      <td>�鳤:
        <input type="text" name="zuz" value="<%=strzuz%>" size="8" /></td>
      <td align="center"><input type="submit" value=" �� �� �� �� " /></td>
    </tr>
  </form>
</table>
<%
End Function

Function TaskList()
	'ȡ�����ݿ��и���ģ�߶�Ӧ�Ķ���������������������TsArray
	Dim TsArray, srow, scol, strhjts
	strhjts=0
	set RS=xjweb.exec("select * from c_tscs order by dmlb",1)
	TsArray = RS.GetRows(,,Array("dmlb","cwts"))
	Rs.close
	Set Rs=Nothing
'-------------
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter
	absPageNum = 0
	strSql = ""
	RecordPerPage=20 		'ÿҳ��ʾ��¼��
	If strLsh <> "" Then
		strSql = " lsh like '%"&strLsh&"%' "
	End If
	If strDwmc <> "" Then
		If strSql <> "" Then
			strSql = " dwmc like '%"&strDwmc&"%' and " & strSql
		Else
			strSql  = " dwmc like '%"&strDwmc&"%' "
		End If
	End If
	If strDmmc <> "" Then
		If strSql <> "" Then
			strSql = " dmmc like '%"&strDmmc&"%' and " & strSql
		Else
			strSql  = " dmmc like '%"&strDmmc&"%' "
		End If
	End If
	If strseb <> "" Then
		If strSql <> "" Then
			strSql = " sbcj like '%"&strseb&"%' and " & strSql
		Else
			strSql  = " sbcj like '%"&strseb&"%'"
		End If
	End If
	If strjcj <> "" Then
		If strSql <> "" Then
			strSql = " jcjxh like '%"&strjcj&"%' and " & strSql
		Else
			strSql  = " jcjxh like '%"&strjcj&"%'"
		End If
	End If
	If strqshu <> "" Then
		If strSql <> "" Then
			strSql = " qs like '%"&strqshu&"%' and " & strSql
		Else
			strSql  = " qs like '%"&strqshu&"%'"
		End If
	End If
	If strzuz <> "" Then
		If strSql <> "" Then
			strSql = " (jgzz like '%"&strzuz&"%' or sjzz like '%"&strzuz&"%' or zz like '%"&strzuz&"%') and " & strSql
		Else
			strSql  = " (jgzz like '%"&strzuz&"%' or sjzz like '%"&strzuz&"%' or zz like '%"&strzuz&"%')"
		End If
	End If
	If strDdh <> "" Then
		If strSql <> "" Then
			strSql = " ddh like '%"&strDdh&"%' and " & strSql
		Else
			strSql  = " ddh like '%"&strDdh&"%' "
		End If
	End If
	If strdxjs <> "" Then
		If strSql <> "" Then
			strSql = " bz like '%"&strdxjs&"%' and " & strSql
		Else
			strSql  = " bz like '%"&strdxjs&"%' "
		End If
	End If

	Select Case strTerm
			Case "no"
				If strSql <> "" Then
					strSql = " isnull(sjjssj) and " & strSql
				Else
					strSql  = " isnull(sjjssj) and not(isnull(ddh))"
				End If
			Case "ok"
				If strSql <> "" Then
					strSql = " not(isnull(sjjssj)) and " & strSql
				Else
					strSql  = " not(isnull(sjjssj)) and not(isnull(ddh))"
				End If
			Case "all"
			Case Else
	End Select

	If Request("strSql") <> "" Then
		strSql = request("strSql")
	End If

	Dim sqlOrder
	sqlOrder = " order by jhjssj desc, lsh desc"
	If LCase(strOrder) = "jhjssj" Then sqlOrder = " order by jhjssj desc, lsh desc"
	If LCase(strOrder) = "ddh" Then sqlOrder = " order by ddh desc, lsh desc"
	If LCase(strOrder) = "lsh" Then sqlOrder = " order by lsh desc, jhjssj desc"
	If LCase(strOrder) = "zz" Then sqlOrder = " order by jgzz desc, sjzz desc, zz desc, lsh desc, jhjssj desc"
	If strSql <> "" Then
		strSql = "select * from [mtask] where "& strSql & sqlOrder
	Else
		strSql = "select * from [mtask] "& sqlOrder
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
	Call TbTopic("����ģ�߳������������")
	iCounter = (absPageNum - 1) * RecordPerPage + 1
%>
<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable" align="center">
  <tr>
    <th class="th" width="25">id
      </td>
    </th>
    <th class="th" width="60">������
      </td>
    </th>
    <th class="th" width="50">��ˮ��
      </td>
    </th>
    <th class="th" width="80">��λ����
      </td>
    </th>
    <th class="th" width="*">��������
      </td>
    </th>
    <th class="th" width="40">�鳤
      </td>
    </th>
    <th class="th" width="80">�ƻ��ṹ
      </td>
    </th>
    <th class="th" width="80">�ƻ�ȫ��
      </td>
    </th>
    <th class="th" width="40">��������
      </td>
    </th>
    <th class="th" width="*">�ṹ
      </td>
    </th>
    <th class="th" width="*">���
      </td>
    </th>
    <th class="th" width="*">���
      </td>
    </th>
    <th class="th" width="*">BOM
      </td>
    </th>
    <th class="th" width="*">����
      </td>
    </th>
  </tr>
  <%
	Dim strzz, strspeed, strbh, strpzxs, strpzdl, strcxfx, strtempa, Arfucai ,strcwts, strtsgj, strtsbh, strtsdq
	For absRecordNum = 1 To RecordPerPage		'��ҳ
	strspeed="" : strbh="" : strpzxs="" : strpzdl="" : strcxfx="" : strtempa=""
	strspeed=Rs("qysd")	: strbh=Rs("xcbh")
	If Rs("zz")<>"" Then
		strzz=Rs("zz")
	ElseIf rs("jgzz")=rs("sjzz") Then
		strzz=Rs("jgzz")
	Else
		strzz=rs("jgzz")&rs("sjzz")
	End If
	strkhlx =NullToNum(Rs("khlb"))
	strcwts=0 : strtsgj=0 : strtsbh=0 : strtsdq=0
  '�ṹ\��ƶ�����ʱ��ͼ��15:54 2007-4-1-������
  '���һ�첻�����ڣ�ʵ��Ҫ���������Ӧ��Ϊdatediff("d", �ƻ���ʼ, �ƻ�����)+1
'���ݿ��˰취�ṹ���ڣ����������-1��/2����˽ṹ���ڣ��������/2
  Dim jgjgsj, sjjgsj, Tmpsj, ijgsj, isj
If rs("jhkssj")<>"" Then
	Tmpsj=rs("jhkssj")
else
	Tmpsj=rs("rwxdsj")
End If
  jgjgsj=datediff("d", Tmpsj, rs("jhjssj"))/2
  sjjgsj=datediff("d", Tmpsj, rs("jhjssj"))+1-jgjgsj
If IsNull(rs("jhjgsj")) Then
	isj=INT(datediff("d", rs("jhkssj"), rs("jhjssj"))/2)
	ijgsj=dateadd("d",isj,rs("jhkssj"))
else
	isj=rs("jhjssj")
	ijgsj=rs("jhjgsj")
End if
%>
  <tr>
    <td class="ctd"><%=iCounter%></td>
    <td class="ctd"><a href="mtask_list.asp?ddh=<%=rs("ddh")%>"><%=rs("ddh")%></a></td>
    <td class="ctd" alt="�Ͳıں�:<%=strbh%>mm&lt;br&gt;���۳��ͷ�϶:<%=strcxfx%>mm&lt;/br&gt;�ο�ƽֱ�γ���:"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("dwmc")%></td>
    <td class="ctd" alt="��������: <%=rs("dmmc")%>"><%=xjweb.StringCut(rs("dmmc"),12)%></td>
    <td class="ctd"><%=strzz%></td>
    <td class="ctd"><%=rs("jhjgsj")%>&nbsp;</td>
    <td class="ctd"><%=rs("jhjssj")%></td>
    <td class="ctd"><%=rs("mjxx") &  rs("rwlr")%></td>
    <%select case rs("mjxx") & rs("rwlr")%>
    <%case "ȫ�����"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <% If not(isnull(rs("gjjgks"))) Then call DisTdjg(rs("gjjgks"),rs("gjjgjs"),ijgsj,rs)%></td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <% If not(isnull(rs("gjsjks"))) Then call DisTdjg(rs("gjsjks"),rs("gjsjjs"),isj,rs)%></td>
    <td class="ctd"><%If not(isnull(rs("mtshr"))) or not(isnull(rs("dxshr"))) Then
      		call distd(rs("mtshks"),rs("mtshjs"),0,rs)
      		call distd(rs("dxshks"),rs("dxshjs"),0,rs)
     		If not(isnull(rs("gjshr"))) Then call distd(rs("gjshks"),rs("gjshjs"),0,rs) End If
     	else
     		call DisTdjg(rs("mtjgshks"),rs("mtjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("dxjgshks"),rs("dxjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("mtsjshks"),rs("mtsjshjs"),isj,rs)
       		call DisTdjg(rs("dxsjshks"),rs("dxsjshjs"),isj,rs)
       		If not(isnull(rs("gjjgr"))) Then call DisTdjg(rs("gjjgshks"),rs("gjjgshjs"),ijgsj,rs) End If
       		If not(isnull(rs("gjsjr"))) Then call DisTdjg(rs("gjsjshks"),rs("gjsjshjs"),isj,rs) End If
      End If%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp;
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "ȫ�׸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
      <%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp;
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "ȫ�׸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp;
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "ģͷ���"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%=jgjgsj%></td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%=sjjgsj%></td>
    <td class="ctd"><%If not(isnull(rs("mtshr"))) Then
      		call distd(rs("mtshks"),rs("mtshjs"),0,rs)
     	else
     		call DisTdjg(rs("mtjgshks"),rs("mtjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("mtsjshks"),rs("mtsjshjs"),isj,rs)
      End If%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "ģͷ����"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "ģͷ����"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "�������"%>
    <td class="ctd"><%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <%=jgjgsj%></td>
    <td class="ctd"><%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <%=sjjgsj%></td>
    <td class="ctd"><%If not(isnull(rs("dxshr"))) Then
      		call distd(rs("dxshks"),rs("dxshjs"),0,rs)
     	else
       		call DisTdjg(rs("dxjgshks"),rs("dxjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("dxsjshks"),rs("dxsjshjs"),isj,rs)
      End If%></td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "���͸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "���͸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%end select
    For srow = 0 To UBound(TsArray, 2)	'���ݵ������ȷ������ʱ��
		If TsArray(0,srow)=Rs("TSLB") Then strcwts=TsArray(1,srow)
	Next
	If Rs("gjzf")>0 Then strtsgj=1	'������1��
	If instr(Rs("xcbh"),"(")>0 Then	'�ں��2.7��1��
		strtsbh=left(Rs("xcbh"),instr(Rs("xcbh"),"(")-1)
	Else
		strtsbh=NullToNum(Rs("xcbh"))
	End If
	If strtsbh>=2.7 Then
		strtsbh=1
	Else
		strtsbh=0
	End If
	If Rs("qs")>1 Then strtsdq=1	'��ǻ��1��
	strcwts=strcwts+strtsgj+strtsbh+strtsdq+strkhlx	'���ݿͻ����ȷ������ʱ��
	strhjts=strhjts+strcwts
    %>
    <td class="ctd" alt="&nbsp;&nbsp;����:<%=Rs("TSLB")%><br>&nbsp;&nbsp;����+<%=strtsgj%><br>&nbsp;&nbsp;�ں�+<%=strtsbh%><br>&nbsp;&nbsp;��ǻ+<%=strtsdq%><br>&nbsp;&nbsp;�ͻ�:<%=strkhlx%><br>----------------<br>�������:<%=strcwts%>��"><%=strcwts%></td>
  </tr>
  <%rs.movenext%>
  <%if rs.eof then%>
  <%exit for%>
  <%end if%>
  <%iCounter=iCounter+1%>
  <%next
  If strDdh<>"" and ChkAble(4) and Session("userGroup")=5 Then
  %>
  <tr>
    <td class="rtd" colspan="7">�ͻ�����:</td>
    <td class="ctd"><select id="khlx" onchange="location.href='?ipage=<%=striPage&strFeedBack%>&action=KhChan&kh='+this.value;">
        <option value=3 <%If strkhlx=3 Then%>selected<%End If%>>A</option>
        <option value=2 <%If strkhlx=2 Then%>selected<%End If%>>B</option>
        <option value=1 <%If strkhlx=1 Then%>selected<%End If%>>C</option>
        <option value=0 <%If strkhlx=0 Then%>selected<%End If%>>D</option>
        <option value=-1 <%If strkhlx=-1 Then%>selected<%End If%>>E</option>
      </select></td>
    <td class="ctd" colspan="4">��ҳ�ϼƳ����������: </td>
    <td class="ctd"><%=strhjts%></td>
  </tr>
  <%End If%>
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

'���Ŀͻ�����
Function KhChan
	strkhlx =Trim(request("kh"))
 	If strFeedBack<>"" Then strFeedBack="?iPage="&striPage&strFeedBack
	strSql="select * from [mtask] where [ddh]='"&strDdh&"'"
	Call xjweb.exec("",-1)
	Rs.open strSql,Conn,1,3
	do while not Rs.eof
		Rs("khlb")=strkhlx
		Rs.update
		Rs.MoveNext
	loop
'	Rs.Close
	Call JsAlert("�ͻ�������ĳɹ�!", "mtask_list.asp"&strFeedBack)
End Function
%>
