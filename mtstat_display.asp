<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="分值统计 → 查看任务分值统计"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="mtstat"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

'定义变量及变量赋值
Dim iyear, imonth, dtstart, dtend, struser, irwzf, iaddfz, ilxrwzf, icount
iyear = request("searchy")
imonth = request("searchm")
struser = request("searchuser")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)

'起始日期
'过渡时期暂时停用
'dtend=cdate(iyear&"年"&imonth&"月10日")
'dtstart=dateadd("m", -1, dtend)		'上个月
'dtstart=dateadd("d", 1, dtstart)		'上个月11日

dtend=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
'dtstart=dateadd("m", -1, dtend)		'上个月
dtstart=cdate(iyear&"年"&imonth&"月1日")


'dtstart=dateadd("d", 6, dtstart)		'上个月11日
'统计人
If struser = "" and chkable(5) Then struser = session("userName")
irwzf=0			'总分
ilxrwzf=0
iaddfz=0		'奖惩分值
icount=1		'工作项目数

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>"  align="center">
  <Tr>
    <Td class=ctd><%Call SearchMantime()%></td>
  </tr>
  <Tr>
    <Td class=ctd height=300><%Call mtstatDisplay()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function SearchMantime()
%>
<table cellpadding=2 cellspacing=0>
  <form action=<%=request.servervariables("script_name")%> method=get>
    <tr>
      <td> 请选择:
        <select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <%for i = year(now) - 3 to year(now)%>
          <option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <%for i = 1 to 12%>
          <option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        月&nbsp;&nbsp;
        <select name="searchuser" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <option value=""></option>
          <%If chkable("1,2,3,4") Then%>
          <%for i = 0 to ubound(c_allstat)%>
          <option value="<%=c_allstat(i)%>" <%If struser = c_allstat(i) Then%>selected<%end If%>><%=c_allstat(i)%></option>
          <%next%>
          <%Else%>
          <option value="<%=session("userName")%>"><%=session("userName")%></option>
          <%end If%>
        </select>
        &nbsp;
        <input type="submit" value=" 选 择 "></td>
    </tr>
  </form>
</table>
<%
End Function

Function mtstatDisplay()
	If struser="" Then
		Call TbTopic("请选择您想查询的人员!")
	Else
		Call TbTopic(struser & " " & formatdatetime(dtstart,1) & " 至 " & formatdatetime(dtend,1) & " 分值统计")
		%>
<table width="80%" cellpadding=2 cellspacing=0 class="xtable"  align="center">
  <tr>
    <th class=th>id</th>
    <th class=th>流水号</th>
    <th class=th>任务内容</th>
    <th class=th>责任人</th>
    <th class=th>完成日期</th>
    <th class=th>任务分值</th>
    <th class=th>系数</th>
    <th class=th>操作</th>
  </tr>
  <%
			call mtask_mt()
			call ftask_mt()
			call total_mt()
		%>
</table>
<%
	End If
End Function

function mtask_mt()		'设计任务分值统计
	strSql="select a.*, b.* ,a.lsh as lsh, a.rwlr as rwlr from [mantime] a, [mtask] b where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 and a.lsh=b.lsh order by jssj desc, a.lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Dim itmpLsh, bJc		'奖惩临时变量
	itmpLsh="" : bJc=True
	Do While Not Rs.eof
	%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd title="<p>订单号: <%=Rs("ddh")%></p>流水号: <%=Rs("lsh")%><br>客户厂家: <%=Rs("dwmc")%><br>断面名称: <%=Rs("dmmc")%><br>"><a href="mtask_display.asp?s_lsh=<%=Rs("lsh")%>"><%=Rs("lsh")%></a></td>
  <td class=ctd><%=Rs("rwlr")%></td>
  <td class=ctd><%=Rs("zrr")%></td>
  <td class=ctd alt="计划结束时间:<%=xjDate(Rs("jhjssj"),1)%>"><%=xjDate(Rs("jssj"),1)%></td>
  <td class=ctd><%=Round(Rs("fz"),1)%></td>
  <td class=ctd><%If Rs("jc")>0 Then Response.Write(Rs("jc")) end If%>
    &nbsp; </td>
  <td class=ctd><%if (InStr(Rs("rwlr"),"调试合格(")>0 and chkable(3)) Then%>
    <input type=button id=<%=Rs("a.id")%> value="修改" onClick="changesjf(this.id)">
    <%End If%>
    &nbsp;</td>
</tr>
<%
		icount = icount + 1
		irwzf=irwzf+Round(Rs("fz"),1)
'		If bJc Then iaddfz=iaddfz+Rs("jc")
		Rs.movenext
	loop
	%>
<tr>
  <td class=rtd colspan=5>任务总分:</td>
  <td class=ctd colspan=3><b><%=irwzf%></b></td>
</tr>
<%
	Rs.close
end function

function ftask_mt()		'零星任务分值统计
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by jssj desc"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.eof or Rs.bof Then
		'response.write("<tr><td class="ctd" colspan=7>没有任何零星任务</td></tr>")
	Else
		Do While Not Rs.eof
		%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd alt="<%=replace(Rs("rwlr"),vbcrlf,"<br>")%>"><%=Rs("id")%></td>
  <td class=ctd><%=Rs("rwlx")%></td>
  <td class=ctd><%=Rs("zrr")%></td>
  <td class=ctd><%=xjDate(Rs("jssj"),1)%></td>
  <td class=ctd><%=Rs("zf")%></td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
</tr>
<%
			icount = icount + 1
			ilxrwzf=ilxrwzf+Rs("zf")
			Rs.movenext
		loop
		%>
<tr>
  <td class=rtd colspan=5>零星任务总分:</td>
  <td class=ctd colspan=3><b><%=ilxrwzf%></b></td>
</tr>
<%
	end If
	Rs.close
end function

function total_mt()		'总分统计
	Dim iTotalFz
	If Fix(ilxrwzf + irwzf + iaddfz)<(ilxrwzf + irwzf + iaddfz) Then
		iTotalFz=Fix(ilxrwzf + irwzf + iaddfz) + 1
	Else
		iTotalFz=Fix(ilxrwzf + irwzf + iaddfz)
	End If
%>
<tr>
  <td class=rtd colspan=5>总分:</td>
  <td class=ctd colspan=3><b><%=iTotalFz%></b></td>
</tr>
<%
end function
%>
<script language="javascript">
function changesjf(arg){
var strsjf
strsjf=showModalDialog("mtstat_c.asp?id="+arg, "", "dialogWidth:280px; dialogHeight:160px; center:yes; help: no; scroll: no; status:no;");
if (strsjf==2){
window.location.reload();
}
}
</script>
