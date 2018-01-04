<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'19:06 2007-4-26-星期四
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="设计任务 → 共挤列表"
strPage="mtask"
xjweb.header()
Call TopTable()

Dim strFeedBack, strTerm,strCjia,iksy,iksm,ijsy,ijsm,ikssj,ijssj
iksy = request("ksy")
iksm = request("ksm")
If iksy = "" Then iksy = year(now)
If iksm = "" Then iksm = month(now)

ijsy = request("jsy")
ijsm = request("jsm")
If ijsy = "" Then ijsy = year(now)
If ijsm = "" Then ijsm = month(now)
ijssj=cdate(ijsy&"年"&ijsm&"月1日")

ijssj=dateadd("m",1,ijssj)
ijssj=dateadd("d",-1,ijssj)
ikssj=cdate(iksy&"年"&iksm&"月1日")
If datediff("d",ikssj,ijssj)<0 Then
	ijssj=cdate(iksy&"年"&iksm&"月1日")
	ijssj=dateadd("m",1,ijssj)
	ijssj=dateadd("d",-1,ijssj)
	ikssj=cdate(ijsy&"年"&ijsm&"月1日")
End If

strTerm=Trim(Request("term"))
strCjia=Trim(Request("cjia"))
strFeedBack="&ksy="&iksy&"&ksm="&iksm&"&jsy="&ijsy&"&jsm="&ijsm&"&term="&strTerm&"&cjia="&strCjia
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()

%>
<table class="xtable" cellspacing="0" cellpadding="2" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd"><%Call SearchInfo()%>
    </td>
  </tr>
  <tr>
    <td class="ctd" height="280"><%Call TaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </td>
  </tr>
</table>
<%
End Sub

Function SearchInfo()
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td>时间范围：
        <select name="ksy" onchange=';'>
          <%for i = year(now) - 12 to year(now) + 1%>
          <option value=<%=i%><%If i = cint(iksy) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="ksm">
          <%for i = 1 to 12%>
          <option value=<%=i%><%If i = cint(iksm) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        月&nbsp;&nbsp;
        &nbsp;--&nbsp;
        <select name="jsy">
          <%for i = year(now) - 12 to year(now) + 1%>
          <option value=<%=i%><%If i = cint(ijsy) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="jsm">
          <%for i = 1 to 12%>
          <option value=<%=i%><%If i = cint(ijsm) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        月&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;条件:
        <select name="term" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;ksy=<%=iksy%>&amp;ksm=<%=iksm%>&amp;jsy=<%= ijsy%>&amp;jsm=<%= ijsm%>&amp;term=&quot; + this.value);'>
          <option value="1" selected="selected">双色共挤</option>
          <option value="2" <%If strterm="2" Then%>selected="selected" <%End If%>>全包覆共挤</option>
          <option value="3" <%If strterm="3" Then%>selected="selected" <%End If%>>软硬前共挤</option>
          <option value="4" <%If strterm="4" Then%>selected="selected" <%End If%>>软硬后共挤</option>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;厂家/端面:
        <input type="text" name="cjia" size="8" value="<%=Trim(Request("cjia"))%>" />
        &nbsp;&nbsp;<input type="submit" value="查 找" />
      </td>
    </tr>
  </form>
</table>
<%
End Function

Function TaskList()
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter
	absPageNum = 0
	strSql = ""
	RecordPerPage=50 		'每页显示记录数

select case strterm
	case "2"
		strSql = " ((gjzf>0 and gjfs=2) or qbfgj<>0)"
	case "3"
		strSql = " ((gjfs=3 and qhgj=1) or qgj<>0)"
	case "4"
		strSql = " ((gjfs=3 and qhgj=2) or hgj<>0)"
	case else
		strSql = " ((gjzf>0 and gjfs=1) or ssgj<>0)"
end select
If strCjia<>"" Then	strSql=" (dwmc like '%"&strCjia&"%' or dmmc like '%"&strCjia&"%') and " & strsql End If

	strSql = "select * from [mtask] where "&strSql&" and datediff('d',jhjssj,'"&ikssj&"')<=0 and datediff('d',jhjssj,'"&ijssj&"')>=0 order by sjjssj desc, lsh desc"
	Set Rs = Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Call xjweb.Exec("",-1)
	Rs.open strSql,Conn,3,3
	If Rs.Eof Or Rs.Bof Then
		Response.Write("没有找到符合条件的模具")
		exit function
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
	Call TbTopic("挤出模具厂设计共挤模具列表")
	iCounter = (absPageNum - 1) * RecordPerPage + 1
%>
<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable">
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
    <th class="th" width="60">单位名称
      </td>
    </th>
    <th class="th" width="60">共挤名称
      </td>
    </th>
    <th class="th" width="60">断面名称
      </td>
    </th>
    <th class="th" width="40">组长
      </td>
    </th>
    <th class="th" width="80">计划完成日期
      </td>
    </th>
    <th class="th" width="*">任务内容
      </td>
    </th>
    <th class="th" width="55">模具结构
      </td>
    </th>
    <th class="th" width="55">模具设计
      </td>
    </th>
    <th class="th" width="55">模具审核
      </td>
    </th>
  </tr>
  <%
	Dim strspeed, strbh, strpzxs, strpzdl, strcxfx, strtempa, Arfucai
	For absRecordNum = 1 To RecordPerPage		'分页
	strspeed="" : strbh="" : strpzxs="" : strpzdl="" : strcxfx="" : strtempa=""
	strspeed=Rs("qysd")	: strbh=Rs("xcbh")
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
  isj=datediff("h", Tmpsj, rs("jhjssj"))+24
	If isj>24 Then
		ijgsj=(isj-24)/2
	else
		isj=24
		ijgsj=24
	End If

Dim ssgjf, qbfgjf, qgjf, hgjf,gjxx
ssgjf=NullToNum(Rs("ssgj"))
qbfgjf=NullToNum(Rs("qbfgj"))
qgjf=NullToNum(Rs("qgj"))
hgjf=NullToNum(Rs("hgj"))
gjxx=""
select case ssgjf&qbfgjf&qgjf&hgjf
Case "0000"			'兼容08版共挤计分模式
	If Rs("gjzf")>0 and Rs("gjfs")=1 Then
 		gjxx="双色共挤"
 	Elseif Rs("gjzf")>0 and Rs("gjfs")=2 Then
  		gjxx="全包覆共挤"
  	Elseif Rs("gjfs")=3 and Rs("qhgj")=1 Then
  		gjxx="软硬前共挤"
  	Elseif Rs("gjfs")=3 and Rs("qhgj")=2 Then
  		gjxx="软硬后共挤"
  	Else
  		gjxx="/"
  	End If
Case Else		'09版共挤计分模式
	If ssgjf<>0 Then gjxx="双色共挤"
	If qbfgjf<>0 Then gjxx=gjxx &" 全包覆共挤"
	If qgjf<>0 Then gjxx=gjxx &" 软硬前共挤"
	If hgjf<>0 Then gjxx=gjxx &" 软硬后共挤"
end select
%>
  <tr>
    <td class="ctd"><%=iCounter%></td>
    <td class="ctd"><%=rs("ddh")%></td>
    <td class="ctd" alt="型材壁厚:<%=strbh%>mm&lt;br&gt;理论成型缝隙:<%=strcxfx%>mm&lt;/br&gt;参考平直段长度:"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("dwmc")%></td>
    <td class="ctd"><%=Trim(gjxx)%></td>
    <td class="ctd"><%=rs("dmmc")%></td>    
    <td class="ctd"><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz"))%>
    </td>
    <%if isnull(rs("sjjssj")) then%>
    <td class="ctd" alt="计划结构完成:<%=Dateadd("h",ijgsj+1,rs("jhkssj"))%><br>整套设计未完成"><%=rs("jhjssj")%></td>
    <%else%>
    <td class="ctd" alt="计划结构完成:<%=Dateadd("h",ijgsj+1,rs("jhkssj"))%><br>实际整套结束日期:<%=rs("sjjssj")%>"><%=rs("jhjssj")%></td>
    <%end if%>
    <td class="ctd"><%=rs("mjxx") &  rs("rwlr")%></td>
    <%select case rs("mjxx") & rs("rwlr")%>
    <%case "全套设计"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%=jgjgsj%>
      <%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <% If not(isnull(rs("gjjgks"))) Then call DisTdjg(rs("gjjgks"),rs("gjjgjs"),ijgsj,rs)%>
    </td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%=sjjgsj%>
      <%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <% If not(isnull(rs("gjsjks"))) Then call DisTdjg(rs("gjsjks"),rs("gjsjjs"),isj,rs)%>
    </td>
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
      End If%>
    </td>
    <%case "全套复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
      <%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <%case "全套复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <%case "模头设计"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%=jgjgsj%> </td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%=sjjgsj%> </td>
    <td class="ctd"><%If not(isnull(rs("mtshr"))) Then
      		call distd(rs("mtshks"),rs("mtshjs"),0,rs)
     	else
     		call DisTdjg(rs("mtjgshks"),rs("mtjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("mtsjshks"),rs("mtsjshjs"),isj,rs)
      End If%>
    </td>
    <%case "模头复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
    </td>
    <%case "模头复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
    </td>
    <%case "定型设计"%>
    <td class="ctd"><%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <%=jgjgsj%> </td>
    <td class="ctd"><%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <%=sjjgsj%> </td>
    <td class="ctd"><%If not(isnull(rs("dxshr"))) Then
      		call distd(rs("dxshks"),rs("dxshjs"),0,rs)
     	else
       		call DisTdjg(rs("dxjgshks"),rs("dxjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("dxsjshks"),rs("dxsjshjs"),isj,rs)
      End If%>
    </td>
    <%case "定型复改"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <%case "定型复查"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <%end select%>
  </tr>
  <%rs.movenext%>
  <%if rs.eof then%>
  <%exit for%>
  <%end if%>
  <%iCounter=iCounter+1%>
  <%next%>
</table>
<table width="98%" cellpadding="2" cellspacing="0" border="0">
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
      </select>
    </td>
  </tr>
</table>
<%
End Function
%>
<script language="javascript">
function checkinf()
	{
		if (trim(document.all.lylr.value)==""){alert("留言内容不能为空！\n");document.all.lylr.focus();document.all.lylr.value="";return false;}
	}
</script>