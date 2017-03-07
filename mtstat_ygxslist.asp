<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="分值统计 → 查看员工系数"					'页面的名称位置( 分值统计 → 查看员工系数)
strPage="mtstat"
xjweb.header()
Call TopTable()

Dim iyear, imonth, dtstart, dtend, irwzf, iaddfz, zcount, icount, ilxrwzf, zrwwcl, zgroup, zdxxs
Dim struser, zrwfz, zrwxs, zzlxs, zgkxs, zbmxs, zjbgz, zjxgz,zyfgz, zbeiz, ygxsRs, m, zbasicwg
zjbgz=0
zjxgz=0
zyfgz=0
zgroup=0
zbasicwg=0
zcount=1
icount=1
iyear = request("searchy")
imonth = request("searchm")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)
dtend=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
dtstart=cdate(iyear&"年"&imonth&"月1日")

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
<table cellpadding=2 cellspacing=0>
  <form action=<%=request.servervariables("script_name")%> method=get>
    <tr>
      <td> 请选择:
        <select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&bybmxs="+this.form.bybmxs.value);'>
          <%for i = year(now) - 3 to year(now)%>
          <option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&bybmxs="+this.form.bybmxs.value);'>
          <%for i = 1 to 12%>
          <option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        月&nbsp;&nbsp;
        <label>本月部门系数：
          <input type="text" name="bybmxs" size="4"  onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&bybmxs="+this.form.bybmxs.value);'>
          &nbsp;&nbsp; </label>
        <label>任务完成率：
          <input type="text" name="rwwcl" size="4">
        </label>
        &nbsp;&nbsp;
        <input type="submit" value=" 确 定 "></td>
    </tr>
  </form>
</table>
<%
End Function

Function YgxsDisplay()		'显示列表
		Call TbTopic("技术部技术人员" & iyear & "年" & imonth & "月考核汇总表")
		If Request("rwwcl")="" Then zrwwcl=1.0 Else zrwwcl=Request("rwwcl") End if
		%>
<table cellpadding=2 cellspacing=0 class="xtable" width="<%=web_info(8)%>">
  <tr>
    <th class=th width="5%">ID</th>
    <th class=th width="8%">人员名单</th>
    <th class=th width="8%">任务分值</th>
    <th class=th width="8%">任务指标</th>
    <th class=th width="12%">任务量考核</th>
    <th class=th width="15%">准时、准确、纪律</th>
    <th class=th width="8%">综合</th>
    <th class=th width="8%">部考系数</th>
    <th class=th width="10%">基本工资</th>
    <th class=th width="10%">绩  效</th>
    <th class=th width="*">应发工资</th>
  </tr>
  <tr>
  	<td colspan="12" class=rtd>本月部门任务完成率=<%=zrwwcl%></td>
  </tr>
</Table>
  <%
		Dim strColor, x
		strColor=-1
		If Request("bybmxs")="" Then zbmxs=1.0 Else zbmxs=Request("bybmxs") End if
		for x = 0 to ubound(c_zypx)
			strSql="select * from [ims_user] where  user_name='"&c_zypx(x)&"'"
			Set ygxsRs=xjweb.Exec(strSql, 1)
			If Not ygxsRs.eof Then
				If zgroup<>ygxsRs("user_group") Then
					strColor=-1*strColor
				End If
				struser=c_zypx(x)
				zgroup=ygxsRs("user_group")
				zbasicwg=ygxsRs("user_basicwage")
			End If
			ygxsRs.close
			zrwfz=0 : zrwxs=0 : zzlxs=0 : zdxxs=0 : zgkxs=0 : zjbgz=0 : zyfgz=0 : zbeiz="" : irwzf=0 : ilxrwzf=0 : iaddfz=0
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
			zrwxs=FormatNumber(zrwxs,2)
%>
			<table cellpadding=2 cellspacing=0  width="<%=web_info(8)%>">
			<tr <%If strColor=1 Then%>bgcolor="#D6D7EF"<%End If%>>
    				<td class=ctd width="5%"><%=zcount%></td>
				<td class=ctd width="8%"><%=struser%></td>
    				<td class=ctd width="8%"><%=zrwfz%>&nbsp;</td>
    				<td class=ctd width="8%"><%=zbasicwg%>&nbsp;</td>
    				<td class=ctd width="12%"><%=zrwxs%>&nbsp;</td>
	    			<td class=ctd width="15%"><%=zzlxs%>&nbsp;</td>
    				<td class=ctd width="8%"><%=zgkxs%>&nbsp;</td>
    				<td class=ctd width="8%"><%=zbmxs%></td>
	    			<td class=ctd width="10%">&nbsp;</td>
    				<td class=ctd width="10%">&nbsp;</td>
    				<td class=ctd width="*">&nbsp;</td>
  			</tr>
  			</Table>
<%
			zcount = zcount + 1
		next
%>
<table cellpadding=2 cellspacing=0 class="xtable" width="<%=web_info(8)%>">
<TR>
	<TD class=rtd colspan=12>The End.</TD>
</TR>
</Table>
<%
End Function

Function YgxsStat()
	strSql="Select * from [ims_user] where [user_name]='"&struser&"'"
	Set Rs=xjweb.Exec(strSql,1)
	Dim tmpCount, tmpGroup, tmpAble, ilxrwzf
	tmpCount=1
	tmpGroup=Rs("user_Group")
	tmpAble=Rs("user_Able")
	Rs.Close

	If InStr("1456",ChkJs(tmpAble))>0 Then		'判断是不是组员或调试员
		'1--任务分值
		strSql="select * from [mantime] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			irwzf=irwzf+Round(Rs("fz"),1)
			Rs.movenext
		Loop
		Rs.close
		'2---零星任务分值
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			ilxrwzf=ilxrwzf+Rs("zf")
			Rs.movenext
		Loop
		Rs.close
		'3---统计总分,
			If Fix(ilxrwzf + irwzf)<(ilxrwzf + irwzf) Then
				zrwfz=Fix(ilxrwzf + irwzf) + 1
			Else
				zrwfz=Fix(ilxrwzf + irwzf)
			End If
	End If
	icount=1
	Select Case ChkJs(tmpAble)
		Case 1	'网管
			kpf(0)=round((zrwfz/zbasicwg * 50),2)
			kpif(0)=statkpfz("大型软件推广应用不及时", 0)
			kpif(1)=statkpfz("技术资料备份不及时", 0)
			kpif(2)=statkpfz("网络权限设定不安全", 0)
			kpf(1)=40+kpif(0)+kpif(1)+kpif(2)
			if kpf(1)<0 Then kpf(1)=0
			kpif(0)=statkpfz("工作态度、劳动纪律扣分", 0)
			kpif(1)=statkpfz("零星任务完成不及时", 0)
			kpf(2)=10 + kpif(0) + kpif(1)
			If kpf(2)<0 Then kpf(2)=0

			for i=1 to 9
				kpzf=kpzf+kpf(i)
			next
			zrwxs=round((kpf(0)/100),2)
			zzlxs=round(kpzf/100,2)
			zgkxs=round(zrwxs+zzlxs,2)
			zyfgz=zjxgz*zgkxs*zbmxs+zjbgz
		Case 6	'调试员
			kpf(0)=round((zrwfz/zbasicwg * 50),2)
			kpif(0)=statkpfz("调试方案问题处理不及时", 0)
			kpif(1)=statkpfz("厂内调试未准时完成", 0)
			kpf(1)=10+kpif(0) + kpif(1)
			if kpf(1)<0 Then kpf(1)=0
			kpif(0)=statkpfz("修理方案原因产生报废", 0)
			kpif(1)=statkpfz("修理方案原因产生返修", 0)
			kpif(2)=statkpfz("设计原因损失超千元", 0)
			kpif(3)=statkpfz("设计原因外部投诉", 0)
			kpf(2)=30+kpif(0)+kpif(1)+kpif(2)+kpif(3)
			if kpf(2)<0 Then kpf(2)=0

			for i=1 to 9
				kpzf=kpzf+kpf(i)
			next
			zrwxs=round((kpf(0)/100),2)
			zzlxs=round(kpzf/100,2)
			zgkxs=round(zrwxs+zzlxs,2)
			zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case Else	'组员
			kpf(0)=round((zrwfz/zbasicwg * 50),2)
			kpif(0)=statkpfz("设计延迟", 0)
			kpf(1)=10+kpif(0)
			if kpf(1)<0 Then kpf(1)=0
  			if ChkJs(tmpAble)=4 Then
				kpif(0)=statkpfz("设计原因产生报废", tmpGroup)
				kpif(1)=statkpfz("设计原因产生返修", tmpGroup)
			Else
				kpif(0)=statkpfz("设计原因产生报废", 0)
				kpif(1)=statkpfz("设计原因产生返修", 0)
			End If
			kpif(2)=statkpfz("设计原因损失超千元", 0)
			kpif(3)=statkpfz("设计原因外部投诉", 0)
			kpf(2)=30+kpif(0)+kpif(1)+kpif(2)+kpif(3)
			if kpf(2)<0 Then kpf(2)=0
			kpif(0)=statkpfz("工作态度、劳动纪律扣分", 0)
			kpif(1)=statkpfz("零星任务完成不及时", 0)
			kpf(3)=10 + kpif(0) + kpif(1)
			If kpf(3)<0 Then kpf(3)=0

			for i=1 to 9
				kpzf=kpzf+kpf(i)
			next
			zrwxs=round((kpf(0)/100),2)
			zzlxs=round(kpzf/100,2)
			zgkxs=round(zrwxs+zzlxs,2)
			zyfgz=zjxgz*zgkxs*zbmxs+zjbgz
	End Select
	Erase kpf
End Function

Function ChkJs(str)
	'str 为权限000001000000000
	ChkJs=0
	If Len(str)<15 Then Exit Function
	dim i
	if Mid(str,8,1)=1 Then	'提升服务技术员优先级
		ChkJs=8
	Else
		For i=1 To Len(str)
			If Mid(str,i,1)=1 Then ChkJs=i : Exit For	'只取每人的最高角色,如你同时是组长和组员,则只取组长
		Next
	End If
End Function

Function statkpfz(kp_item, i)
	Dim ZzSql, ZzRs
	statkpfz=0
	Dim tmpRs
	Select Case i
		Case 0		'对组员进行统计
			strSql="select ([kp_uprice]*[kp_mul]) as kp_f from [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'对组长进行统计
			strSql="select [kp_lsh],max([kp_uprice]*[kp_mul]*0.5) as kp_f from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 group by [kp_lsh]"
			ZzSql="select ([kp_uprice]*[kp_mul]*0.5) as kp_f from [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End Select

	Set tmpRs=xjweb.Exec(strSql, 1)
	do while not tmpRs.eof
		statkpfz=statkpfz + tmpRs("kp_f")
		tmpRs.movenext
	loop
	tmpRs.close
	set tmprs=nothing

	If i>0 Then
		Set ZzRs=xjweb.Exec(ZzSql, 1)
		Do While not ZzRs.eof
			statkpfz=statkpfz + ZzRs("kp_f")
			ZzRs.movenext
		loop
		ZzRs.close
		set ZzRs=nothing
	End If
	statkpfz=round(statkpfz,2)
End Function
%>