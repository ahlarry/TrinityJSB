<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="升级数据"					'页面的名称位置( 任务书管理 → 添加任务书)yutg
strPage=""
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, strlsh
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder


Dim iyear, imonth, dtstart, dtend, irwzf, iaddfz, zcount, icount, ilxrwzf, zrwwcl, zgroup, zdxxs
Dim struser, zrwfz, zrwxs, zzlxs, zgkxs, zbmxs, zjbgz, zjxgz,zyfgz, zbeiz, ygxsRs, m, xrwfz
zjbgz=0
zjxgz=0
zyfgz=0
zgroup=0
xrwfz=0
zcount=1
icount=1
dtstart=cdate("2017年2月1日")
dtend=dateadd("m",1,dtstart)
dtend=dateadd("d",-1,dtend)

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
    <Td class=ctd height=300>
      <%
     	 Call YgxsDisplay()
      %>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function UpFz()
	Dim n, slsh, oldxs, newxs
'	slsh="12109"
'	slsh=split(slsh,",")
'	newxs=1
'	for n=0 to ubound(slsh)
'		oldxs=1
'		strSql="select * from [mtask] where lsh='"&slsh(n)&"'"
'		Call xjweb.exec("",-1)
'		Rs.open strSql,Conn,1,3
'		If Not Rs.eof Then
'			oldxs=rs("fzxs")
'			rs("fzxs")=newxs
'		End If
'		Rs.update
'		Rs.close

'		strSql="select * from [mantime] where (Instr(rwlr,'结构')>0 or Instr(rwlr,'设计')>0 or Instr(rwlr,'BOM')>0) and lsh='"&slsh(n)&"'"
'		Call xjweb.exec("",-1)
'		Rs.open strSql,Conn,1,3
'			Do while not Rs.eof
'				Rs("fz")=Round(Rs("fz")*newxs/oldxs,1)
'				Rs.update
'				Rs.movenext
'			Loop
'		Rs.close

'	next
end function

Function YgxsDisplay()		'显示列表
		Call TbTopic("技术部技术人员" & dtstart & "-" & dtend & "月考核汇总表")
		If Request("rwwcl")="" Then zrwwcl=1.0 Else zrwwcl=Request("rwwcl") End if
		%>
<table cellpadding=2 cellspacing=0 class="xtable" width="<%=web_info(8)%>">
  <tr>
    <th class=th width="15%">ID</th>
    <th class=th width="15%">人员名单</th>
    <th class=th width="30%">任务分值</th>
    <th class=th width="20%">任务x1.2</th>
    <th class=th >任务x1.3</th>
  </tr>
</Table>
  <%
		Dim strColor, x, newxs
		strColor=-1
		zbmxs=1.0
		for x = 0 to ubound(c_zypx)
			strSql="select * from [ims_user] where  user_name='"&c_zypx(x)&"'"
			Set ygxsRs=xjweb.Exec(strSql, 1)
			If Not ygxsRs.eof Then
				If zgroup<>ygxsRs("user_group") Then
					strColor=-1*strColor
				End If
				struser=c_zypx(x)
				zgroup=ygxsRs("user_group")
			End If
			ygxsRs.close
			Call YgxsStat()
'%>
			<table cellpadding=2 cellspacing=0  width="<%=web_info(8)%>">
			<tr <%If strColor=1 Then%>bgcolor="#D6D7EF"<%End If%>>
    				<td class=ctd width="15%"><%=zcount%></td>
				<td class=ctd width="15%"><%=struser%></td>
    				<td class=ctd width="30%"><%=zrwfz%>&nbsp;</td>
<%
		newxs=1
		if struser="江烈环" Then zrwfz=zrwfz-100
		if zrwfz>400 Then
			newxs=1.3
		else if zrwfz>300 Then
				newxs=1.2
			End If
		End If
		if newxs>1 Then
			call xsup(newxs)
			Call YgxsStat()
		End If
'%>
    				<td class=ctd width="20%"><%if newxs=1.2 Then Response.write(zrwfz)%>&nbsp;</td>
    				<td class=ctd><%if newxs=1.3 Then Response.write(zrwfz)%>&nbsp;</td>
  			</tr>
  			</Table>
<%
			zcount = zcount + 1
		next
'%>
<table cellpadding=2 cellspacing=0 class="xtable" width="<%=web_info(8)%>">
<TR>
	<TD class=rtd colspan=12>The End.</TD>
</TR>
</Table>
<%
End Function

Function YgxsStat()
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

	strSql="Select * from [ims_user] where [user_name]='"&struser&"'"
	Set Rs=xjweb.Exec(strSql,1)
	Dim tmpCount, tmpGroup, tmpAble, ilxrwzf
	tmpCount=1
	tmpGroup=Rs("user_Group")
	tmpAble=Rs("user_Able")
	Rs.Close

	If InStr("456",ChkJs(tmpAble))>0 Then		'判断是不是组员或调试员
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
End Function

Function xsup(newxs)
	strSql="Select * from [ims_user] where [user_name]='"&struser&"'"
	Set Rs=xjweb.Exec(strSql,1)
	Dim tmpCount, tmpGroup, tmpAble, ilxrwzf
	tmpCount=1
	tmpAble=Rs("user_Able")
	Rs.Close

	If InStr("456",ChkJs(tmpAble))>0 Then		'判断是不是组员或调试员
		'1--任务分值
		strSql="select * from [mantime] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Do While Not Rs.eof
			Rs("fz")=Round(Rs("fz")*newxs,1)
			Rs.update
			Rs.movenext
		Loop
		Rs.close

		'2---零星任务分值
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 and zrr<>'江烈环'"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Do While Not Rs.eof
			Rs("zf")=Round(Rs("zf")*newxs,1)
			Rs.update
			Rs.movenext
		Loop
		Rs.close
	End If
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
%>
