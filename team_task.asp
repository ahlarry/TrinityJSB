<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->

<%
'10:37 2011-12-07
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="设计任务 → 本组任务"
strPage="mtstat"
xjweb.header()
Call TopTable()

Dim iyear, imonth, dtstart, dtend, struser, irwzf, iaddfz, ilxrwzf, icount
iyear = request("searchy")
imonth = request("searchm")
struser = request("searchuser")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)

dtend=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
dtstart=cdate(iyear&"年"&imonth&"月1日")

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
<table class="xtable" cellspacing="0" cellpadding="2" width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchMantime()%></td>
  </tr>
  <tr>
    <td class="ctd" height="280"><%Call TaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%></td>
  </tr>
</table>
<%
End Sub

Function TaskList()
	Dim iGroup, tmpSql, tmpRs
	Call TbTopic(struser & "组" & imonth & "月任务定额")
%>
	<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable"  align="center">
  <tr>
    <th class="th" width="20">id
      </td>
    </th>
    <th class="th">流水号
      </td>
    <th class="th" width="120">任务内容
      </td>
    </th>
    <th class="th">责任人
      </td>
    </th>
    <th class="th" width="120">角色
      </td>
    </th>
    <th class="th" width="*">定额断面
      </td>
    </th>
    <th class="th" width="100">常规定额
      </td>
    </th>
    <th class="th" width="100">本套定额
      </td>
    </th>
  </tr>
  <%
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&struser&"'"
		Set tmpRs=xjweb.Exec(tmpSql,1)
		If Not(tmpRs.Eof Or tmpRs.Bof) Then
			iGroup=tmpRs("user_group")
		Else
			iGroup=0
		End If
	tmpRs.Close
				  
  	strSql="select a.*, b.* ,a.lsh as lsh, a.xz as xz,b.rwlr as rwlr from [reward] a, [mtask] b where xz="&iGroup&" and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 and a.lsh=b.lsh order by jssj desc, a.lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Dim itmpLsh, bJc		'奖惩临时变量
	itmpLsh="" : bJc=True
	Do While Not Rs.eof
	%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd ><a href="mtask_display.asp?s_lsh=<%=Rs("lsh")%>"><%=Rs("lsh")%></a></td>
  <td class=ctd title=<%=Rs("demt")&","&Rs("dedx")%>>
  		<%If Not(IsNull(Rs("mtrw"))) Then Response.Write("模头"&Rs("mtrw")) End If%>&nbsp;
  		<%If Not(IsNull(Rs("dxrw"))) Then Response.Write("定型"&Rs("dxrw")) End If%>
  </td>
  <td class=ctd><%=Rs("zrr")%></td>
  <td class=ctd><%=Rs("js")%></td>
  <td class=ctd><%if Rs("dedm")<>"" Then Response.Write(Rs("dedm")) else Response.Write(Rs("ckdm"))%></td>
  <td class=ctd title=""><%=Round(Rs("fz"),1)%></td>
  <td class=ctd><%=Round(Rs("fz"),1)%></td>
</tr>	
<%
		icount = icount + 1
		irwzf=irwzf+Round(Rs("fz"),1)
		Rs.movenext
	loop
	%>
<tr>
  <td class=rtd colspan=7>任务总分:</td>
  <td class=ctd><b><%=irwzf%></b></td>
</tr>
<%
	Rs.close
  %>
</table>
<%
End Function

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
        月&nbsp;
        <select name="searchuser" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <option value=""></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value="<%=c_allzz(i)%>" <%If struser = c_allzz(i) Then%>selected<%end If%>><%=c_allzz(i)%></option>
          <%next%>
        </select>
        <input type="submit" value=" 选 择 "></td>
    </tr>
  </form>
</table>
<%
End Function
%>
