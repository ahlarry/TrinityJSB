<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="分值统计 → 分项统计"				
strPage="mtstat"
xjweb.header()
Call TopTable()

Dim isjfz, iaddfz, zcount, idbsh, iksy, iksm, ijsy, ijsm, ikssj, ijssj, struser, idbsj, ilxxl, ilxrw, ygxsRs, ishfz

zcount=1
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

'定义考评用的变量
	Dim kpf(30), kpif(10), ics(10), kpzf, kpxr
	kpxr=Array("")

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchMantime()%></td>
  </tr>
 </Table>
<%Call YgxsDisplay()
      Response.Write(XjLine(10,"100%",""))
End Sub

Function SearchMantime()
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
        月
        &nbsp;&nbsp;<input type="submit" value="统  计" />
      </td>
    </tr>
  </form>
</table>
<%
End Function

Function YgxsDisplay()		'显示列表
		Call TbTopic("技术人员" & ikssj & "至" & ijssj & "分项统计表")
		%>
<table cellpadding=2 cellspacing=0 class="xtable" width="<%=web_info(8)%>">
  <tr>
    <th class=th rowspan="2" width="5%">ID</th>
    <th class=th rowspan="2" width="10%">姓名</th>
    <th class=th colspan="3" width="25%">模具</th>
    <th class=th colspan="3" width="25%">技术代表</th>
    <th class=th rowspan="2" width="12%">零星修理</th>
    <th class=th rowspan="2" width="12%">零星任务</th>
    <th class=th rowspan="2">总分</th>
  </tr>
  <tr>
    <th class=th>设计</th>
    <th class=th>审核</th>
    <th class=th>合计</th>
    <th class=th>设计</th>
    <th class=th>审核</th>
    <th class=th>合计</th>
  </tr>
  <%
		Dim strColor, x
		for x = 0 to ubound(c_zypx)
			strSql="select * from [ims_user] where  user_name='"&c_zypx(x)&"'"
			Set ygxsRs=xjweb.Exec(strSql, 1)
			If Not ygxsRs.eof Then 
				struser=c_zypx(x)		
			End If
			ygxsRs.close
			kpzf=0
			for i=0 to 29
				kpf(i)=0
			next
			for i=0 to 9
				kpif(i)=0
			next
			for i=0 to 9
				ics(i)=0
			next
			Call YgxsStat()
%>
			<tr>
    				<td class=ctd ><%=zcount%></td>
				<td class=ctd ><%=struser%></td>
    				<td class=ctd ><%=isjfz%></td>
    				<td class=ctd ><%=ishfz%></td>
    				<td class=ctd bgcolor="#D6D7EF"><%=isjfz+ishfz%></td>
	    			<td class=ctd ><%=idbsj%></td>
    				<td class=ctd ><%=idbsh%></td>
    				<td class=ctd bgcolor="#D6D7EF"><%=idbsj+idbsh%></td>
    				<td class=ctd ><%=ilxxl%></td>
    				<td class=ctd ><%=ilxrw%></td>
    				<td class=ctd bgcolor="#D6D7EF"><%=Round(isjfz+ishfz+idbsj+idbsh+ilxxl+ilxrw,1)%></td>
  			</tr>
<%
			zcount = zcount + 1
		next
%>
<TR>
	<TD class=rtd colspan=12>The End.</TD>
</TR>
</Table>
<%
End Function

Function YgxsStat()
		isjfz=0 : ishfz=0 : idbsj=0 : idbsh=0 : ilxxl=0 : ilxrw=0
		'1--模具设计分值
		strSql="select * from [mantime] where zrr='"&struser&"' and datediff('d',jssj,'"&ikssj&"')<=0 and datediff('d',jssj,'"&ijssj&"')>=0 and (Right(rwlr,len('结构'))='结构' or Right(rwlr,len('设计'))='设计') "
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			isjfz=isjfz+Rs("fz")
			Rs.movenext
		Loop
		isjfz=Round(isjfz,1)
		Rs.close
		'2--模具审核分值
		strSql="select * from [mantime] where zrr='"&struser&"' and datediff('d',jssj,'"&ikssj&"')<=0 and datediff('d',jssj,'"&ijssj&"')>=0 and Right(rwlr,len('确认'))='确认' "
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			ishfz=ishfz+Rs("fz")
			Rs.movenext
		Loop
		ishfz=Round(ishfz,1)
		Rs.close
		'3--技术代表设计分值
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&ikssj&"')<=0 and datediff('d',jssj,'"&ijssj&"')>=0 and rwlx='技术代表设计' "
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			idbsj=idbsj+Rs("zf")
			Rs.movenext
		Loop
		idbsj=Round(idbsj,1)
		Rs.close
		'4--技术代表审核分值
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&ikssj&"')<=0 and datediff('d',jssj,'"&ijssj&"')>=0 and rwlx='技术代表审核' "
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			idbsh=idbsh+Rs("zf")
			Rs.movenext
		Loop
		idbsh=Round(idbsh,1)
		Rs.close
		'4--零星修理分值
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&ikssj&"')<=0 and datediff('d',jssj,'"&ijssj&"')>=0 and rwlx='零星修理' "
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			ilxxl=ilxxl+Rs("zf")
			Rs.movenext
		Loop
		ilxxl=Round(ilxxl,1)
		Rs.close
		'5--零星任务分值
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&ikssj&"')<=0 and datediff('d',jssj,'"&ijssj&"')>=0 and rwlx='零星任务' "
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			ilxrw=ilxrw+Rs("zf")
			Rs.movenext
		Loop
		ilxrw=Round(ilxrw,1)
		Rs.close
End Function
%>