<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(6)
CurPage="模具调试 → 填写调试信息"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="mtest"
Call FileInc(0, "js/mtest.js")
xjweb.header()
Call TopTable()

	Dim tscs, pscs, s_lsh
	'tscs--调试次数 pscs--评审次数
	tscs=0 : pscs = 0 : s_lsh=request("s_lsh")
	tscs=xjweb.RsCount("ts_tsxx where lsh='" & s_lsh & "' and not ps")
	pscs=xjweb.Rscount("ts_tsxx where lsh='" & s_lsh & "' and ps")

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
			<%Call mtestAdd()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function mtestAdd()
	If s_lsh="" Then Call TbTopic("请输入添加调试信息模具的流水号!") : Exit Function
	Dim regEx '建立变量。
	Set regEx = New RegExp ' 建立正则表达式。
	regEx.Pattern = "^[a-zA-z][0-9]+" ' 设置模式。
	regEx.IgnoreCase = False ' 设置是否区分大小写。
	If regEx.Test(s_lsh)  Then
		strSql="select a.*, b.*,a.lsh as lsh from  [ts_mould] a, [ftask] b where a.lsh='"&s_lsh&"' and a.lsh=b.xlxh and isnull(tsjssj)"
	Else
		strSql="select a.*, b.*,a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh='"&s_lsh&"' and a.lsh=b.lsh and isnull(tsjssj)"
	End If	
	
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof Or rs.bof Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务不存在、已完成或调试手册没有完成!","mtest_add.asp")
	ElseIf Not IsNull(Rs("tsjssj")) Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务模具调试工作已经结束!")
	Else
		Call mould_inf(rs,regEx.Test(s_lsh))
		Response.Write(xjLine(10, "100%", ""))
		If tscs > 0 And tscs mod 5 = 0 And int(tscs/5) > pscs Then
			Call mtestps_add(rs)
		Else
			call mtest_add(rs)
		End If
		Call PreNext(s_lsh)
		Response.write(XjLine(10, "100%", ""))				
	End If
End Function

Function mtest_add(rs)
%>
	<%Call TbTopic("添加 <span alt=""流水号"">" &rs("lsh")&"</span> 模具第 "&(tscs+ 1) &" 次调试信息")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
	<form id=frm_mtestadd name=frm_mtestadd action=mtest_indb.asp?action=add method=post onSubmit='return tscheckinf();'>

	<tr>
		<th class=rtd height=25 width="20%">项目名称</td>
		<th class=ctd width="*">项目内容</td>
	</tr>
	<tr>
		<td class=rtd>调试原因</td>
		<td class=ltd><textarea name="tsyy" cols="95" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd>调试内容</td>
		<td class=ltd><textarea name="tslr" cols="95" rows="7"></textarea></td>
	</tr>
	<tr><td class=ctd colspan=2><input type=submit value=" · 添 加 · "></td></tr>
	<input type="hidden" name="lsh" value=<%=rs("lsh")%>>
	<input type="hidden" name="tsps" value=false>
	</form>

	</table>
<%
end function		'mtest_add()

Function mtestps_add(rs)
%>
	<%Call TbTopic("<font style=color:#ff0000>添加流水号 " &rs("lsh")&" 模具第 "&(pscs+1)&" 次评审记录</font>")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
	<form id=frm_mtestpsadd name=frm_mtestpsadd action=mtest_indb.asp?action=add method=post onSubmit='return tspscheckinf();'>

	<tr>
		<th class=rtd height=25 width="20%">项目名称</td>
		<th class=ctd width="*">项目内容</td>
	</tr>
	<tr>
		<td class=rtd>评审内容</td>
		<td class=ltd><textarea name="tslr" cols="95" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd>评审人</td>
		<td class=ltd><textarea name="tsyy" cols="95" rows="3"></textarea></td>
	</tr>
	<tr><td class=ctd colspan=2><input type=submit value=" · 添 加 · "></td></tr>
	<input type="hidden" name="lsh" value=<%=rs("lsh")%>>
	<input type="hidden" name="tsps" value=true>
	</form>
	</table>
<%
End Function

Function rwlr_change(i)
         dim mystr,mystr1,mystr2
		 mystr=rs("rwlr")
			 If Instr(mystr,"||")>0 Then
			     mystr=split(mystr,"||")
			     If i > ubound(mystr) Then
			     	mystr1=""
			     	rwlr_change=mystr1
			     else
	   		 		 mystr1=mystr(i)
					 mystr1=split(mystr1,":")
					 rwlr_change=mystr1(1)
				 End If

			 else
			    mystr=split(mystr,chr(10))
	   		 	mystr1=mystr(i)
	   		 	If Instr(mystr1,"：")>0 Then
					mystr1=split(mystr1,"：")
					rwlr_change=mystr1(1)
				else
					rwlr_change=mystr1
				End If
			 End If
End Function

Function mould_inf(Rs,xl)
	Dim strrwlr, strddh, strlsh, strdwmc, strmjxx, strdmmc
	strrwlr="" : strddh="" : strlsh="" : strdwmc="" : strmjxx="" : strdmmc=""
	If xl Then
		strddh=rs("xldh")
		strlsh=rwlr_change(2)
		strdwmc=rwlr_change(0)
		strmjxx=rwlr_change(6)
		strdmmc=rwlr_change(1)
	Else
		strddh=rs("ddh")
		strlsh=rs("lsh")
		strdwmc=rs("dwmc")
		strmjxx=rs("mjxx")
		strdmmc=rs("dmmc")
	End If	
%>
<%Call TbTopic("流水号 "&Rs("lsh")&" 模具信息")%>
<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
  <tr>
    <td class=th width="10%">订单号</td>
    <td class=th width="*">断面名称</td>
    <td class=th width="10%">流水号</td>
    <td class=th width="10%">单位名称</td>
    <td class=th width="10%">调试类型</td>
    <td class=th width="15%">调试开始</td>
    <td class=th width="15%">最近调试</td>
    <td class=th width="10%">调试次数</td>
  </tr>
  <tr>
    <td class=ctd><%=strddh%></td>
    <td class=ctd><%=strdmmc%></td>
    <td class=ctd><a href="mtask_display.asp?s_lsh=<%=strlsh%>"><%=strlsh%></a></td>
    <td class=ctd><%=strdwmc%></td>
    <td class=ctd><%=strmjxx%></td>
    <td class=ctd><%=xjDate(Rs("tskssj"),1)%>&nbsp;</td>
    <td class=ctd alt="<%If isnull(Rs("tsjssj")) Then%>正在调试<%Else%>调试结束<%End If%>"><%=xjDate(Rs("tsgxsj"),1)%>&nbsp;</td>
    <td class=ctd><%=xjweb.RsCount("ts_tsxx where lsh='"&Rs("lsh")&"' and not(ps)")%></td>
  </tr>
</table>
<%
End Function

Function mtest_display(lsh)
	Dim prs, itscs, ipscs
	strSql="select * from [ts_tsxx] where lsh='"&lsh&"' order by id desc"
	itscs=xjweb.rscount("[ts_tsxx] where lsh='"&lsh&"' and not(ps)")
	ipscs=xjweb.rscount("[ts_tsxx] where lsh='"&lsh&"' and ps")
	Set prs = xjweb.Exec(strSql, 1)
	If Prs.Eof Or Prs.Bof Then Prs.Close : Set Prs=Nothing : Call TbTopic("暂时没有任何调试信息!") : Exit Function
	Call TbTopic("流水号 " &lsh&" 模具调试信息列表")
%>
<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
  <%
	do while not prs.eof
		If prs("ps") Then
	%>
  <tr bgcolor=#dddddd>
    <td class=ctd width="10%" rowspan="3">第 <b><%=ipscs%></b> 次<br>
      评审</td>
    <td class=rtd width="10%">评审内容:</td>
    <td class=ltd width="*"><%=xjweb.htmltocode(prs("tslr"))%></td>
  </tr>
  <tr bgcolor=#dddddd>
    <td class=rtd>评审人:</td>
    <td class=ltd><%=xjweb.htmltocode(prs("tsyy"))%></td>
  </tr>
  <form action="mtest_indb.asp?action=delete" method=post onsubmit="return confirm('确认删除吗?');">
    <tr bgcolor=#dddddd>
      <td class=rtd colspan="2">签写:<%=prs("tsr")%> 日期:<%=prs("tssj")%>
        <%If chkable(6) and prs("tsr")=Session("userName") Then%>
        &nbsp;<a href="mtest_change.asp?id=<%=prs("id")%>&cs=<%=ipscs%>&s_lsh=<%=lsh%>&ps=true">编缉</a>&nbsp;
        <%End If%>
        <%If chkable(1) Then%>
        <input type="submit" value=" 删除 ">
        <input type="hidden" name=id value="<%=prs("id")%>">
        <input type="hidden" name="lsh" value="<%=prs("lsh")%>">
        <%End If%></td>
    </tr>
  </form>
  <%
			ipscs=ipscs-1
		Else
	%>
  <tr>
    <td class=ctd width="10%" rowspan="3">第 <b><%=itscs%></b> 次</td>
    <td class=rtd width="10%">调试原因:</td>
    <td class=ltd width="*"><%=xjweb.htmltocode(prs("tsyy"))%></td>
  </tr>
  <tr>
    <td class=rtd>调试内容:</td>
    <td class=ltd><%=xjweb.htmltocode(prs("tslr"))%></td>
  </tr>
  <form action="mtest_indb.asp?action=delete" method=post onsubmit="return confirm('确认删除吗?');">
    <tr>
      <td class=rtd colspan="2">调试:<%=prs("tsr")%> 日期:<%=prs("tssj")%>
        <%If chkable(6) and prs("tsr")=Session("userName") Then%>
        &nbsp;<a href="mtest_change.asp?id=<%=prs("id")%>&cs=<%=itscs%>&s_lsh=<%=lsh%>&ps=false">编缉</a>&nbsp;
        <%End If%>
        <%If chkable(1) Then%>
        <input type="submit" value=" 删除 ">
        <input type="hidden" name="id" value="<%=prs("id")%>">
        <input type="hidden" name="lsh" value="<%=prs("lsh")%>">
        <%End If%></td>
    </tr>
  </form>
  <%
			itscs=itscs-1
		End If
		prs.movenext
	loop
	prs.close
	Set prs = nothing
	%>
</table>
<%
End Function

Function PreNext(ilsh)
Dim strOrder,strPre,strNext,TmpSql,Trs,Tsj
strOrder=Trim(Request("order")) : strPre="" : strNext="" : Tsj=""

	TmpSql="select * from [ts_mould] where  lsh = '" &ilsh& "'"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	Tsj=Trs("tsgxsj")
	Trs.close
	Set Trs = nothing

If strOrder="tsgxsj" Then
	TmpSql="select * from [ts_mould] where datediff('s',tsgxsj,'"&Tsj&"')>0 and isnull(tsjssj) order by tsgxsj desc,lsh desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then 
		strPre="Beg"
	Else
		strPre=Trs("lsh")
	End  If
	TmpSql="select * from [ts_mould] where datediff('s',tsgxsj,'"&Tsj&"')<0 and isnull(tsjssj) order by tsgxsj,lsh desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then 
		strNext="End"
	Else
		strNext=Trs("lsh")
	End  If
Else
	TmpSql="select a.*, b.*,a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh < '"&ilsh&"' and isnull(tsjssj) and a.lsh=b.lsh order by a.lsh desc,tsgxsj desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then 
		strPre="Beg"
	Else
		strPre=Trs("lsh")
	End  If
	TmpSql="select a.*, b.*,a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh > '"&ilsh&"' and isnull(tsjssj) and a.lsh=b.lsh order by a.lsh,tsgxsj desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then 
		strNext="End"
	Else
		strNext=Trs("lsh")
	End  If
End If
Trs.close
Set Trs = nothing
%>
<table cellspacing=0 cellpadding=3 width="95%" align="center">
  <tr>
    <td width="20%">
    <%If strPre="Beg" Then
    	Response.write("")
    else%>
   		<a href=mtest_add.asp?s_lsh=<%=strPre%>&order=<%=strOrder%>><strong>上一个：<%=strPre%></strong></a>
   	<%End If%>
    </td>
    <td width="*" align="center">排序:
      <select name="order" onchange='location.href("<%=Request.servervariables("script_name")%>?s_lsh=<%=ilsh%>&order=" + this.value);'>
        <option value="" selected="selected">流 水 号</option>
        <option value="tsgxsj" <%If strOrder="tsgxsj" Then%>selected<%End If%>>最近调试</option>
      </select>请注意流水号中的"<font size="4" color="#ff0000"><strong>C</strong></font>"</td>
    <td width="20%" align="right">
    <%If strNext="End" Then
    	Response.write("")
    else%>    
    	<a href=mtest_add.asp?s_lsh=<%=strNext%>&order=<%=strOrder%>><strong>下一个：<%=strNext%></strong></a>
    <%End If%>
    </td>
  </tr>
</Table>
<%End Function%>
