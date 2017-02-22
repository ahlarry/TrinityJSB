<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'14:32 2007-1-22-星期一
Call ChkPageAble(0)
'Call ChkDepart("技术部")
CurPage="设计任务 → 任务流程"
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
      <td>&nbsp;&nbsp;订单号:
        <input type="text" name="ddh" value="<%=strDdh%>" size="10" /></td>
      <td>流水号:
        <input tabindex="1" type="text" name="lsh" size="8" value="<%=strLsh%>" /></td>
      <td>单位:
        <input type="text" name="dwmc" value="<%=strDwmc%>" size="8" /></td>
      <td>断面:
        <input type="text" name="dmmc" value="<%=strDmmc%>" size="8" /></td>
      <td>排序:
        <select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;lsh=<%= strlsh%>&amp;ddh=<%=strddh%>&amp;dwmc=<%=strdwmc%>&amp;term=<%=strterm%>&amp;seb=<%=strseb%>&amp;jcj=<%=strjcj%>&amp;qshu=<%=strqshu%>&amp;order=&quot; + this.value);'>
          <%If strOrder = "lsh" Then%>
          <option value="jhjssj">计划完成时间</option>
          <option value="ddh">订单号</option>
          <option value="lsh" selected="selected">流水号</option>
          <option value="zz">组长</option>
          <%ElseIf strOrder="ddh" Then%>
          <option value="jhjssj">计划完成时间</option>
          <option value="ddh" selected="selected">订单号</option>
          <option value="lsh">流水号</option>
          <option value="zz">组长</option>
          <%ElseIf strOrder="zz" Then%>
          <option value="jhjssj">计划完成时间</option>
          <option value="ddh">订单号</option>
          <option value="lsh">流水号</option>
          <option value="zz" selected="selected">组长</option>
          <%Else%>
          <option value="jhjssj" selected="selected">计划完成时间</option>
          <option value="ddh">订单号</option>
          <option value="lsh">流水号</option>
          <option value="zz">组长</option>
          <%End If%>
        </select></td>
      <td>条件:
        <select name="term" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;lsh=<%= strlsh%>&amp;ddh=<%=strddh%>&amp;dwmc=<%=strdwmc%>&amp;seb=<%=strseb%>&amp;jcj=<%=strjcj%>&amp;qshu=<%=strqshu%>&amp;order=<%=strorder%>&amp;term=&quot; + this.value);'>
          <%select case strterm%>
          <%case "no"%>
          <option value="no" selected="selected">未完成</option>
          <option value="ok">已完成</option>
          <option value="all">全部</option>
          <%case "ok"%>
          <option value="no">未完成</option>
          <option value="ok" selected="selected">已完成</option>
          <option value="all">全部</option>
          <%case "all"%>
          <option value="no">未完成</option>
          <option value="ok">已完成</option>
          <option value="all" selected="selected">全部</option>
          <%case else%>
          <option value="no">未完成</option>
          <option value="ok">已完成</option>
          <option value="all" selected="selected">全部</option>
          <%end select%>
        </select></td>
    </tr>
    <tr>
      <td>&nbsp;&nbsp;设备厂:
        <input type="text" name="seb" value="<%=strseb%>" size="10" /></td>
      <td>挤出机:
        <input type="text" name="jcj" value="<%=strjcj%>" size="8" /></td>
      <td>腔数:
        <input type="text" name="qshu" value="<%=strqshu%>" size="8" /></td>
      <td>备注:
        <input type="text" name="dxjs" value="<%=strdxjs%>" size="8" /></td>
      <td>组长:
        <input type="text" name="zuz" value="<%=strzuz%>" size="8" /></td>
      <td align="center"><input type="submit" value=" · 查 找 · " /></td>
    </tr>
  </form>
</table>
<%
End Function

Function TaskList()
	'取出数据库中各类模具对应的额定厂外调试天数，存入数组TsArray
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
	RecordPerPage=20 		'每页显示记录数
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
		Call JsAlert("指定条件的任务书不存在,请放宽筛选条件!","")
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
	Call CutLine()		'显示图例
	Call TbTopic("挤出模具厂设计任务流程")
	iCounter = (absPageNum - 1) * RecordPerPage + 1
%>
<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable" align="center">
  <tr>
    <th class="th" width="25">id
      </td>
    </th>
    <th class="th" width="60">订单号
      </td>
    </th>
    <th class="th" width="50">流水号
      </td>
    </th>
    <th class="th" width="80">单位名称
      </td>
    </th>
    <th class="th" width="*">断面名称
      </td>
    </th>
    <th class="th" width="40">组长
      </td>
    </th>
    <th class="th" width="80">计划结构
      </td>
    </th>
    <th class="th" width="80">计划全套
      </td>
    </th>
    <th class="th" width="40">任务内容
      </td>
    </th>
    <th class="th" width="*">结构
      </td>
    </th>
    <th class="th" width="*">设计
      </td>
    </th>
    <th class="th" width="*">审核
      </td>
    </th>
    <th class="th" width="*">BOM
      </td>
    </th>
    <th class="th" width="*">厂外
      </td>
    </th>
  </tr>
  <%
	Dim strzz, strspeed, strbh, strpzxs, strpzdl, strcxfx, strtempa, Arfucai ,strcwts, strtsgj, strtsbh, strtsdq
	For absRecordNum = 1 To RecordPerPage		'分页
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
  '结构\设计独立计时的图例15:54 2007-4-1-星期日
  '最后一天不算延期，实际要求设计周期应该为datediff("d", 计划开始, 计划结束)+1
'根据考核办法结构周期＝（设计周期-1）/2，因此结构周期＝设计周期/2
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
    <td class="ctd" alt="型材壁厚:<%=strbh%>mm&lt;br&gt;理论成型缝隙:<%=strcxfx%>mm&lt;/br&gt;参考平直段长度:"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("dwmc")%></td>
    <td class="ctd" alt="断面名称: <%=rs("dmmc")%>"><%=xjweb.StringCut(rs("dmmc"),12)%></td>
    <td class="ctd"><%=strzz%></td>
    <td class="ctd"><%=rs("jhjgsj")%>&nbsp;</td>
    <td class="ctd"><%=rs("jhjssj")%></td>
    <td class="ctd"><%=rs("mjxx") &  rs("rwlr")%></td>
    <%select case rs("mjxx") & rs("rwlr")%>
    <%case "全套设计"%>
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
    <%case "全套复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
      <%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp;
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "全套复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp;
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "模头设计"%>
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
    <%case "模头复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "模头复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "定型设计"%>
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
    <%case "定型复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%case "定型复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
      &nbsp; </td>
    <%end select
    For srow = 0 To UBound(TsArray, 2)	'根据调试类别确定厂外时间
		If TsArray(0,srow)=Rs("TSLB") Then strcwts=TsArray(1,srow)
	Next
	If Rs("gjzf")>0 Then strtsgj=1	'共挤加1天
	If instr(Rs("xcbh"),"(")>0 Then	'壁厚≥2.7加1天
		strtsbh=left(Rs("xcbh"),instr(Rs("xcbh"),"(")-1)
	Else
		strtsbh=NullToNum(Rs("xcbh"))
	End If
	If strtsbh>=2.7 Then
		strtsbh=1
	Else
		strtsbh=0
	End If
	If Rs("qs")>1 Then strtsdq=1	'多腔加1天
	strcwts=strcwts+strtsgj+strtsbh+strtsdq+strkhlx	'根据客户类别确定厂外时间
	strhjts=strhjts+strcwts
    %>
    <td class="ctd" alt="&nbsp;&nbsp;调试:<%=Rs("TSLB")%><br>&nbsp;&nbsp;共挤+<%=strtsgj%><br>&nbsp;&nbsp;壁厚+<%=strtsbh%><br>&nbsp;&nbsp;多腔+<%=strtsdq%><br>&nbsp;&nbsp;客户:<%=strkhlx%><br>----------------<br>厂外调试:<%=strcwts%>天"><%=strcwts%></td>
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
    <td class="rtd" colspan="7">客户类型:</td>
    <td class="ctd"><select id="khlx" onchange="location.href='?ipage=<%=striPage&strFeedBack%>&action=KhChan&kh='+this.value;">
        <option value=3 <%If strkhlx=3 Then%>selected<%End If%>>A</option>
        <option value=2 <%If strkhlx=2 Then%>selected<%End If%>>B</option>
        <option value=1 <%If strkhlx=1 Then%>selected<%End If%>>C</option>
        <option value=0 <%If strkhlx=0 Then%>selected<%End If%>>D</option>
        <option value=-1 <%If strkhlx=-1 Then%>selected<%End If%>>E</option>
      </select></td>
    <td class="ctd" colspan="4">本页合计厂外调试天数: </td>
    <td class="ctd"><%=strhjts%></td>
  </tr>
  <%End If%>
</table>
<table width="98%" cellpadding="2" cellspacing="0" border="0" align="center">
  <tr>
    <td align="left"> 符合条件共 <%=Rs.RecordCount%> 个&nbsp;&nbsp;
      每页 <%=Rs.PageSize%> 个&nbsp;&nbsp;
      共 <%=Rs.PageCount%> 页&nbsp;&nbsp;
      当前为第 <%=absPageNum%> 页 </td>
    <td align="right"> 【
      <%
				If absPageNum>1 Then
					Response.Write("<a href="&request.servervariables("script_name")&"?ipage="&(abspagenum-1)&strFeedBack&" title='上一页'> ←</a>&nbsp;&nbsp;")
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
					Response.Write("&nbsp;<a href="&Request.ServerVariables("SCRIPT_NAME")&"?iPage="&(abspagenum+1)&strFeedBack&" title='下一页'> → </a>&nbsp;")
				End If
			%>
      】
      跳转到:
      <select name="ipage" onchange='location.href(&quot;<%=Request.ServerVariables("SCRIPT_NAME")%>?ipage=&quot; + this.value +&quot;<%=strFeedback%>&quot;);'>
        <%For i=1 To Rs.PageCount%>
        <%If i=absPageNum Then%>
        <option value="<%=i%>" selected="selected">第 <%=i%> 页</option>
        <%Else%>
        <option value="<%=i%>">第 <%=i%> 页</option>
        <%End If%>
        <%Next%>
      </select></td>
  </tr>
</table>
<%
End Function

'更改客户分类
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
	Call JsAlert("客户分类更改成功!", "mtask_list.asp"&strFeedBack)
End Function
%>
