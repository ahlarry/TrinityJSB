<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'11:32 2007-4-10-星期二
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="分值统计 → 查看考评分值统计"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="mtstat"
xjweb.header()
Call TopTable()

'定义变量及变量赋值
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

'定义考评用的变量
	Dim kpf(30), kpif(10), ics(10), kpzf
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
  <Tr>
    <Td class=ctd height=300><%Call ygkpstatDisplay()%>
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

Function ygkpstatDisplay()
	If struser="" Then Call TbTopic("请选择您想查询的人员!") : Exit Function
	strSql="Select * from [ims_user] where [user_name]='"&struser&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then TbTopic("请重新选择查询人员!") : Rs.Close : Exit Function
	Dim tmpGroup, tmpAble
	tmpGroup=Rs("user_Group")
	tmpAble=Rs("user_Able")
	Rs.Close

	Dim iTotalFz, tmpCount			'定义总分的变量
	iTotalFz=0 : tmpCount=1
	If InStr("5689",ChkJs(tmpAble))>0 Then		'判断是不是组员或调试员
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
		'3---统计总分
		If Fix(ilxrwzf + irwzf)<(ilxrwzf + irwzf) Then
			iTotalFz=Fix(ilxrwzf + irwzf) + 1
		Else
			iTotalFz=Fix(ilxrwzf + irwzf)
		End If
	else
		If InStr("4",ChkJs(tmpAble))>0 Then		'判断是不是组长
			'1--统计小组成员人数
			strSql="Select * from [ims_user] where [user_group]="&tmpGroup
			Call xjweb.Exec("",-1)
			Set Rs=Server.CreateObject("ADODB.RECORDSET")
			Rs.open strSql,Conn,1,3
			tmpCount=Rs.RecordCount
			Rs.Close
			'2--任务分值
			strSql="select * from [mantime] where [xz]="&tmpGroup&" and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
			Set Rs=xjweb.Exec(strSql, 1)
			Do While Not Rs.eof
				irwzf=irwzf+Round(Rs("fz"),1)
				Rs.movenext
			Loop
			Rs.close
			'3---零星任务分值
			strSql="select * from [ftask] where [xz]="&tmpGroup&" and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
			Set Rs=xjweb.Exec(strSql, 1)
			Do While Not Rs.eof
				ilxrwzf=ilxrwzf+Rs("zf")
				Rs.movenext
			Loop
			Rs.close
			'4---统计总分
			iTotalFz=Round((ilxrwzf + irwzf)/tmpCount,1)
		End If
	End If

	icount=1
	Call TbTopic(struser & " " & formatdatetime(dtstart,1) & " 至 " & formatdatetime(dtend,1) & " 考评统计")
		%>
<table width="90%" cellpadding=2 cellspacing=0 class="xtable"  align="center">
<tr>
  <th class=th>id</th>
  <th class=th>考评项目</th>
  <th class=th>考评指标</th>
  <th class=th>单元分(分)</th>
  <th class=th>次分(分)</th>
  <th class=th>单位</th>
  <th class=th>总次数</th>
  <th class=th>考评应得分</th>
  <th class=th>考评实际分</th>
</tr>
<%
	Select Case ChkJs(tmpAble)
		Case 4	'组长
			%>
<%icount=1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>任务量</td>
  <td class=ctd>小组人平均分值300分</td>
  <td class=ctd>50.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <%
					If iTotalFz<300 Then
						kpf(0)=round((iTotalFz/300 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-300)/300*50*1.25)),1)
					End If
				%>
  <td class=ctd alt="<%="任务:" & iTotalFz & "分"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>设计完成</td>
  <td class=ltd>提前</td>
  <td class=ctd rowspan=2>10</td>
  <td class=ctd>1</td>
  <td class=ctd>分/天</td>
  <%
					ics(0)=statkpcs("提前", "", 0)
					ics(1)=statkpcs("延迟", "", 0)

					kpif(0)=statkpfz("提前", 0)
					kpif(1)=statkpfz("延迟", 0)
					kpf(1)=kpif(0) + kpif(1)
					if kpf(1)<-10 Then kpf(1)=-10
					if kpf(1)>10 Then kpf(1)=10
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>延迟</td>
  <td class=ctd>2</td>
  <td class=ctd>分/天</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=7>设计质量</td>
  <td class=ltd>模具设计和任务书不符</td>
  <td class=ctd rowspan=7>20</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("模具设计和任务书不符", "", 0)
					ics(1)=statkpcs("设计原因产生返工", "审核", 0)+statkpcs("设计原因产生返工", "设计", tmpGroup)
					ics(2)=statkpcs("设计原因产生返修", "审核", 0)+statkpcs("设计原因产生返修", "设计", tmpGroup)
					ics(3)=statkpcs("设计原因产生报废", "审核", 0)+statkpcs("设计原因产生报废", "设计", tmpGroup)
					ics(4)=statkpcs("厂内调试少于额定次数", "", 0)
					ics(5)=statkpcs("厂内调试多于额定次数", "", 0)
					ics(6)=statkpcs("不同批次相似型材类似问题重复发生", "", 0)

					kpif(0)=statkpfz("模具设计和任务书不符", 0)
					kpif(1)=statkpfz("设计原因产生返工", 0) - ics(1)
					kpif(2)=statkpfz("设计原因产生返修", 0) - ics(2)*2
					kpif(3)=statkpfz("设计原因产生报废", 0) - ics(3)*4
					kpif(4)=statkpfz("厂内调试少于额定次数", 0)
					kpif(5)=statkpfz("厂内调试多于额定次数", 0)
					kpif(6)=statkpfz("不同批次相似型材类似问题重复发生", 0)

					kpf(2)=kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4) + kpif(5) + kpif(6)
					if kpf(2)<-20 Then kpf(2)=-20
					if kpf(2)>20 Then kpf(2)=20
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=7><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>设计原因产生返工</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>设计原因产生返修</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>设计原因产生报废</td>
  <td class=ctd>4</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>厂内调试少于额定次数</td>
  <td class=ctd>0.4</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>厂内调试多于额定次数</td>
  <td class=ctd>0.4</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(5)%></td>
  <td class=ctd><%=kpif(5)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>不同批次相似型材类似问题重复发生</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(6)%></td>
  <td class=ctd><%=kpif(6)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
    <%
				'将质量分随任务分变化而变化,即质量分=任务分/300*30,但不大于30
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				If iTotalFz>300 Then
					kpzf=kpzf+30
				else
					kpzf=round(kpzf+(iTotalFz/300 * 30),1)
				End If
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%icount=1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=6>内部管理</td>
  <td class=ltd>设计规范维护及时性</td>
  <td class=ctd rowspan=6>20</td>
  <td class=ctd>5</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("设计规范维护及时性", "", 0)
					ics(1)=statkpcs("项目实施计划未按时完成", "", 0)
					ics(2)=statkpcs("组内技术代表维护不及时", "", 0)
					ics(3)=statkpcs("标准结构物料未充分利用", "", 0)
					ics(4)=statkpcs("现场管理、卫生清洁", "", 0)
					ics(5)=statkpcs("不服从分配", "", 0)

					kpif(0)=statkpfz("设计规范维护及时性", 0)
					kpif(1)=statkpfz("项目实施计划未按时完成", 0)
					kpif(2)=statkpfz("组内技术代表维护不及时", 0)
					kpif(3)=statkpfz("标准结构物料未充分利用", 0)
					kpif(4)=statkpfz("现场管理、卫生清洁", 0)
					kpif(5)=statkpfz("不服从分配", 0)

					kpf(3)=kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4) + kpif(5)
					if kpf(3)<-20 Then kpf(3)=-20
					if kpf(3)>20 Then kpf(3)=20
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=6><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>项目实施计划未按时完成</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>组内技术代表维护不及时</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>标准结构物料未充分利用</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>现场管理、卫生清洁</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ltd>不服从分配</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(5)%></td>
  <td class=ctd><%=kpif(5)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <td class=ctd><%=20+kpf(3)%></td>
</tr>
<%
		Case 6	'调试员
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>任务量</td>
  <td class=ctd>每月分值(总量300分)</td>
  <td class=ctd>50.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <%
					If iTotalFz<300 Then
						kpf(0)=round((iTotalFz/300 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-300)/300*50*1.25)),1)
					End If
				%>
  <td class=ctd alt="<%="任务:" & iTotalFz & "分"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>任务完成</td>
  <td class=ctd>延迟</td>
  <td class=ctd rowspan=2>10.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>分/次(多人平分)</td>
  <%
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)

					kpif(0)=statkpfz("延迟", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("提前", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>提前</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>分/次(多人平分)</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>调试制度</td>
  <td class=ctd>5次以上修理未评审</td>
  <td class=ctd rowspan=4>8.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("5次以上修理未评审", "", 0)
					ics(1)=statkpcs("修理方案下发不及时", "", 0)
					ics(2)=statkpcs("修理情报录入不及时", "", 0)
					ics(3)=statkpcs("修理图纸签署、更改不完善", "", 0)

					kpif(0)=statkpfz("5次以上修理未评审", 0)
					kpif(1)=statkpfz("修理方案下发不及时", 0)
					kpif(2)=statkpfz("修理情报录入不及时", 0)
					kpif(3)=statkpfz("修理图纸签署、更改不完善", 0)

					kpf(2)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(2)<-8 Then kpf(2)=-8
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=4><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>修理方案下发不及时</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>修理情报录入不及时</td>
  <td class=ctd>1.5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>修理图纸签署、更改不完善</td>
  <td class=ctd>1.5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>调试技能</td>
  <td class=ctd>修理图未及时存档</td>
  <td class=ctd rowspan=3>20.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("修理图未及时存档", "", 0)
					ics(1)=statkpcs("修理方案原因产生返修", "", 0)
					ics(1)=ics(1)+statkpcs("设计原因产生返修", "", 0)
					ics(2)=statkpcs("修理方案原因产生报废", "", 0)
					ics(2)=ics(2)+statkpcs("设计原因产生报废", "", 0)

					kpif(0)=statkpfz("修理图未及时存档", 0)
					kpif(1)=statkpfz("修理方案原因产生返修", 0)
					kpif(1)=kpif(1)+statkpfz("设计原因产生返修", 0)
					kpif(2)=statkpfz("修理方案原因产生报废", 0)
					kpif(2)=kpif(2)+statkpfz("设计原因产生报废", 0)

					kpf(3)=kpif(0)+kpif(1)+kpif(2)
					If kpf(3)<-20 Then kpf(3)=-20
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>修理方案原因产生返修</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>修理方案原因产生报废</td>
  <td class=ctd>4.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>设计验证</td>
  <td class=ctd>调试超过额定最大次数</td>
  <td class=ctd rowspan=2>&nbsp;</td>
  <td class=ctd>0.15</td>
  <td class=ctd>×模具分值×超出次数</td>
  <%
					ics(0)=statkpcs("调试超过额定最大次数", "", 0)
					ics(1)=statkpcs("调试少于额定最小次数", "", 0)
					ics(2)=statkpcs("模具调试未合格数", "", 0)

					kpif(0)=statkpfz("调试超过额定最大次数", 0)
					kpif(1)=statkpfz("调试少于额定最小次数", 0)
					kpif(2)=statkpfz("模具调试未合格数", 0)

					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-6 Then kpf(4)=-6
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>调试少于额定最小次数</td>
  <td class=ctd>0.15</td>
  <td class=ctd>×模具分值×少于次数</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>模具调试未合格数</td>
  <td class=ctd>6.0</td>
  <td class=ctd>3.0</td>
  <td class=ctd>分/副(平分)</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>工作纪律</td>
  <td class=ctd>上班做与工作无关的事</td>
  <td class=ctd rowspan=3>2.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("值班离岗,闲聊", "", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", "", 0)

					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("值班离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)

					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(6)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>值班离岗,闲聊</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>桌面不洁,下班机器未关、门未锁</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>工作态度</td>
  <td class=ctd>不服从分配</td>
  <td class=ctd rowspan=2>4.0</td>
  <td class=ctd>4.0</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("不服从分配", "", 0)
					ics(1)=statkpcs("消极怠工", "", 0)

					kpif(0)=statkpfz("不服从分配", 0)
					kpif(1)=statkpfz("消极怠工", 0)

					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(7)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>消极怠工</td>
  <td class=ctd>4.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				'将质量分由默认的50改为随任务分变化而变化,即质量分=任务分/300*50,但不大于50
				If iTotalFz>300 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(iTotalFz/300 * 50),1)
				End If
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
		Case 5	'组员
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>任务量</td>
  <td class=ctd>每月分值(总量300分)</td>
  <td class=ctd>50.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <%
					If iTotalFz<300 Then
						kpf(0)=round((iTotalFz/300 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-300)/300*50*1.25)),1)
					End If
				%>
  <td class=ctd alt="<%="任务:" & iTotalFz & "分"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>任务完成</td>
  <td class=ctd>延迟</td>
  <td class=ctd rowspan=2>10.0</td>
  <td class=ctd>2.5</td>
  <td class=ctd>分/天</td>
  <%
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)

					kpif(0)=statkpfz("延迟", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("提前", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>提前</td>
  <td class=ctd>1.5</td>
  <td class=ctd>分/天</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=5>设计质量</td>
  <td class=ctd>设计原因产生报废</td>
  <td class=ctd rowspan=5>20.0</td>
  <td class=ctd>4.0</td>
  <td class=ctd>分/件</td>
  <%
					ics(0)=statkpcs("设计原因产生报废", "", 0)
					ics(1)=statkpcs("设计原因产生返修", "", 0)
					ics(2)=statkpcs("设计原因产生返工", "", 0)
					ics(3)=statkpcs("厂内调试少于额定次数", "", 0)
					ics(4)=statkpcs("厂内调试多于额定次数", "", 0)

					kpif(0)=statkpfz("设计原因产生报废", 0)
					kpif(1)=statkpfz("设计原因产生返修", 0)
					kpif(2)=statkpfz("设计原因产生返工", 0)
					kpif(3)=statkpfz("厂内调试少于额定次数", 0)
					kpif(4)=statkpfz("厂内调试多于额定次数", 0)

					kpf(2)=kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)
					If kpf(2)<-20 Then kpf(2)=-20
					If kpf(2)>20 Then kpf(2)=20
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=5><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>设计原因产生返修</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/件</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>设计原因产生返工</td>
  <td class=ctd>1.0</td>
  <td class=ctd>分/件</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>厂内调试少于额定次数</td>
  <td class=ctd>1.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>厂内调试多于额定次数</td>
  <td class=ctd>1.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				'将质量分由默认的50改为随任务分变化而变化,即质量分=任务分/300*50,但不大于50
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				If iTotalFz>300 Then
					kpzf=kpzf+30
				else
					kpzf=round(kpzf+(iTotalFz/300 * 30),1)
				End If
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%icount=1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>技术提升</td>
  <td class=ctd>提出改进建议取得成效</td>
  <td class=ctd rowspan=2>&nbsp;</td>
  <td class=ctd>1.0~5.0</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("提出改进建议取得成效", "", 0)
					ics(1)=statkpcs("提出合理化建议并被采纳", "", 0)

					kpif(0)=statkpfz("提出改进建议取得成效", 0)
					kpif(1)=statkpfz("提出合理化建议并被采纳", 0)

					kpf(3)=kpif(0) + kpif(1)
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>提出合理化建议并被采纳</td>
  <td class=ctd>1.0~5.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>纪律态度</td>
  <td class=ctd>上班做与工作无关的事</td>
  <td class=ctd rowspan=3>20.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("不服从分配", "", 0)
					ics(2)=statkpcs("主动承担较难任务", "", 0)

					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("不服从分配", 0)
					kpif(2)=statkpfz("主动承担较难任务", 0)

					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-20 Then kpf(4)=-20
					If kpf(4)>20 Then kpf(4)=20
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>不服从分配</td>
  <td class=ctd>2.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>主动承担较难任务</td>
  <td class=ctd>3.0~5.0</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				kpzf=20
				for i=3 to 4
					kpzf=kpzf+kpf(i)
				next
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
		Case 8		'工艺技术员
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>任务量</td>
  	<td class=ctd>每月分值(总量400分)</td>
  <td class=ctd>50</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <%
					If iTotalFz<400 Then
						kpf(0)=round((iTotalFz/400 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-400)/400*50*1.25)),1)
					End If
				%>
  <td class=ctd alt="<%="任务书:" & iTotalFz & "分"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>任务完成</td>
  <td class=ctd>延迟</td>
  <td class=ctd rowspan=2>10</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次(</td>
  <%
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)

					kpif(0)=statkpfz("延迟", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("提前", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>提前</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>项目完成</td>
  <td class=ctd>新品开发月计划任务未完成</td>
  <td class=ctd rowspan=3>5</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("新品开发月计划任务未完成", "", 0)
					ics(1)=statkpcs("新品开发月、周未及时上报", "", 0)
					ics(2)=statkpcs("新品开发实施计划维护情况差", "", 0)

					kpif(0)=statkpfz("新品开发月计划任务未完成", 0)
					kpif(1)=statkpfz("新品开发月、周未及时上报", 0)
					kpif(2)=statkpfz("新品开发实施计划维护情况差", 0)

					kpf(2)=kpif(0) + kpif(1) + kpif(2)
					If kpf(2)<-5 Then kpf(2)=-5
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>新品开发月、周未及时上报</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>新品开发实施计划维护情况差</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=10>工艺质量</td>
  <td class=ctd>工艺文件更改不完善未造成质量问题</td>
  <td class=ctd rowspan=10>20</td>
  <td class=ctd>0.5</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("工艺文件更改不完善未造成质量问题", "", 0)
					ics(1)=statkpcs("所编制的作业指导书经论证后未及时修改", "", 0)
					ics(2)=statkpcs("工艺文件签署不完整", "", 0)
					ics(3)=statkpcs("因签署日期不实，导致质量争议", "", 0)
					ics(4)=statkpcs("因工艺错误造成返修（漏、错工艺）", "", 0)
					ics(5)=statkpcs("因工艺错误造成报废", "", 0)
					ics(6)=statkpcs("未按规范执行", "", 0)
					ics(7)=statkpcs("因工艺造成车间订单不能下达", "", 0)
					ics(8)=statkpcs("相同错误重复出现", "", 0)
					ics(9)=statkpcs("设计结构、材料、热处理等明显错误未及时反映", "", 0)

					kpif(0)=statkpfz("工艺文件更改不完善未造成质量问题", 0)
					kpif(1)=statkpfz("所编制的作业指导书经论证后未及时修改", 0)
					kpif(2)=statkpfz("工艺文件签署不完整", 0)
					kpif(3)=statkpfz("因签署日期不实，导致质量争议", 0)
					kpif(4)=statkpfz("因工艺错误造成返修（漏、错工艺）", 0)
					kpif(5)=statkpfz("因工艺错误造成报废", 0)
					kpif(6)=statkpfz("未按规范执行", 0)
					kpif(7)=statkpfz("因工艺造成车间订单不能下达", 0)
					kpif(8)=statkpfz("相同错误重复出现", 0)
					kpif(9)=statkpfz("设计结构、材料、热处理等明显错误未及时反映", 0)

					If kpif(8)>4 Then kpif(8)=10
					kpf(3)=kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)+kpif(5)+kpif(6)+kpif(7)+kpif(8)+kpif(9)
					If kpf(3)<-20 Then kpf(3)=-20
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=10><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>所编制的作业指导书经论证后未及时修改</td>
  <td class=ctd>3</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>工艺文件签署不完整</td>
  <td class=ctd>0.5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>因签署日期不实，导致质量争议</td>
  <td class=ctd>3</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>因工艺错误造成返修（漏、错工艺）</td>
  <td class=ctd>1.5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>因工艺错误造成报废</td>
  <td class=ctd>3</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(5)%></td>
  <td class=ctd><%=kpif(5)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>未按规范执行</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(6)%></td>
  <td class=ctd><%=kpif(6)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>因工艺造成车间订单不能下达</td>
  <td class=ctd>0.5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(7)%></td>
  <td class=ctd><%=kpif(7)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>相同错误重复出现</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(8)%></td>
  <td class=ctd><%=kpif(8)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>设计结构、材料、热处理等明显错误未及时反映</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(9)%></td>
  <td class=ctd><%=kpif(9)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>技能水平</td>
  <td class=ctd>主动解决非工艺因素造成的加工问题</td>
  <td class=ctd rowspan=4>5</td>
  <td class=ctd>+</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("主动解决非工艺因素造成的加工问题", "", 0)
					ics(1)=statkpcs("主动提出加工工艺改进方案并被采纳", "", 0)
					ics(2)=statkpcs("制作专用夹具成功并在计划期内推行", "", 0)
					ics(3)=statkpcs("新进厂技术员技能考核不合格", "", 0)

					kpif(0)=statkpfz("主动解决非工艺因素造成的加工问题", 0)
					kpif(1)=statkpfz("主动提出加工工艺改进方案并被采纳", 0)
					kpif(2)=statkpfz("制作专用夹具成功并在计划期内推行", 0)
					kpif(3)=statkpfz("新进厂技术员技能考核不合格", 0)

					kpf(4)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(4)<-5 Then kpf(4)=-5
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=4><%=kpf(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>主动提出加工工艺改进方案并被采纳</td>
  <td class=ctd>+</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>制作专用夹具成功并在计划期内推行</td>
  <td class=ctd>+</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>新进厂技术员技能考核不合格</td>
  <td class=ctd>5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>工作纪律</td>
  <td class=ctd>上班做与工作无关的事</td>
  <td class=ctd rowspan=3>5</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("值班离岗,闲聊", "", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", "", 0)

					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("值班离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)

					kpf(5)=kpif(0) + kpif(1) + kpif(2)
					If kpf(5)<-5 Then kpf(5)=-5
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(5)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>值班离岗,闲聊</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>桌面不洁,下班机器未关、门未锁</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>工作态度</td>
  <td class=ctd>不服从分配</td>
  <td class=ctd rowspan=4>5</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("不服从分配", "", 0)
					ics(1)=statkpcs("消极怠工", "", 0)
					ics(2)=statkpcs("处理问题不及时，且无正当理由", "", 0)
					ics(3)=statkpcs("主动承担较难任务并积极完成", "", 0)

					kpif(0)=statkpfz("不服从分配", 0)
					kpif(1)=statkpfz("消极怠工", 0)
					kpif(2)=statkpfz("处理问题不及时，且无正当理由", 0)
					kpif(3)=statkpfz("主动承担较难任务并积极完成", 0)

					kpf(6)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(6)<-5 Then kpf(6)=-5
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=4><%=kpf(6)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>消极怠工</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>处理问题不及时，且无正当理由</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>主动承担较难任务并积极完成</td>
  <td class=ctd>3</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				'将质量分由默认的50改为随任务分变化而变化,即质量分=任务分/400*50,但不大于50
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If iTotalFz>400 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(iTotalFz/400 * 50),1)
				End If
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
		Case 9			'编程技术员
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>任务量</td>
  	<td class=ctd>每月分值(总量400分)</td>
  <td class=ctd>50.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <%
					If iTotalFz<400 Then
						kpf(0)=round((iTotalFz/400 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-400)/400*50*1.25)),1)
					End If
				%>
  <td class=ctd alt="<%="任务书:" & iTotalFz & "分"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>任务完成</td>
  <td class=ctd>延迟</td>
  <td class=ctd rowspan=2>10</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)

					kpif(0)=statkpfz("延迟", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("提前", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>提前</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次(多人平分)</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=5>程序质量</td>
  <td class=ctd>因程序错误出现返修</td>
  <td class=ctd rowspan=5>20</td>
  <td class=ctd>1.5</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("因程序错误出现返修", "", 0)
					ics(1)=statkpcs("相同错误重复出现2次以上", "", 0)
					ics(2)=statkpcs("因程序错误出现报废", "", 0)
					ics(3)=statkpcs("加工者自检发现程序错误", "", 0)
					ics(4)=statkpcs("编程发现图纸有设计问题", "", 0)

					kpif(0)=statkpfz("因程序错误出现返修", 0)
					kpif(1)=statkpfz("相同错误重复出现2次以上", 0)
					kpif(2)=statkpfz("因程序错误出现报废", 0)
					kpif(3)=statkpfz("加工者自检发现程序错误", 0)
					kpif(4)=statkpfz("编程发现图纸有设计问题", 0)

					kpf(2)=kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)
					If kpf(2)<-20 Then kpf(2)=-20
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=5><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>相同错误重复出现2次以上</td>
  <td class=ctd>3</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>因程序错误出现报废</td>
  <td class=ctd>3</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>加工者自检发现程序错误</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>编程发现图纸有设计问题</td>
  <td class=ctd>0.5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>技能水平</td>
  <td class=ctd>解决非常规数控设备加工问题</td>
  <td class=ctd rowspan=4>10</td>
  <td class=ctd>+</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("解决非常规数控设备加工问题", "", 0)
					ics(1)=statkpcs("提出落料改进并被采用", "", 0)
					ics(2)=statkpcs("出现落料错误", "", 0)
					ics(3)=statkpcs("新进厂技术员技能考核不合格", "", 0)

					kpif(0)=statkpfz("解决非常规数控设备加工问题", 0)
					kpif(1)=statkpfz("提出落料改进并被采用", 0)
					kpif(2)=statkpfz("出现落料错误", 0)
					kpif(3)=statkpfz("新进厂技术员技能考核不合格", 0)

					kpf(3)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(3)<-10 Then kpf(3)=-10
					If kpf(3)>10 Then kpf(3)=10
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=4><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>提出落料改进并被采用</td>
  <td class=ctd>+</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>出现落料错误</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>新进厂技术员技能考核不合格</td>
  <td class=ctd>5</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>工作纪律</td>
  <td class=ctd>上班做与工作无关的事</td>
  <td class=ctd rowspan=3>5</td>
  <td class=ctd>1</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("无正当理由离岗,闲聊", "", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", "", 0)

					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("无正当理由离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)

					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-2 Then kpf(4)=-2
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>无正当理由离岗,闲聊</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>桌面不洁,下班机器未关、门未锁</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>工作态度</td>
  <td class=ctd>不服从分配</td>
  <td class=ctd rowspan=3>5</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <%
					ics(0)=statkpcs("不服从分配", "", 0)
					ics(1)=statkpcs("处理问题不及时", "", 0)
					ics(2)=statkpcs("消极怠工", "", 0)

					kpif(0)=statkpfz("不服从分配", 0)
					kpif(1)=statkpfz("处理问题不及时", 0)
					kpif(2)=statkpfz("消极怠工", 0)

					kpf(5)=kpif(0) + kpif(1) + kpif(2)
					If kpf(5)<-5 Then kpf(5)=-5
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(5)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>处理问题不及时</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>消极怠工</td>
  <td class=ctd>2</td>
  <td class=ctd>分/次</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				'将质量分由默认的50改为随任务分变化而变化,即质量分=任务分/400*50,但不大于50
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If iTotalFz>400 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(iTotalFz/400 * 50),1)
				End If
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
	End Select
	Erase kpf
End Function

Function ChkJs(str)
	'str 为权限000001000000000
	ChkJs=0
	'If IsDebug Then ChkAble=True : Exit Function
	If Len(str)<15 Then Exit Function
	dim i
	For i=1 To Len(str)
		If Mid(str,i,1)=1 Then ChkJs=i : Exit For	'只取每人的最高角色,如你同时是组长和组员,则只取组长
	Next
End Function

Function statkpjfz(kp_item,zrrjs,i)
	Dim PjCs,tmpRs
	statkpjfz=0
	strSql="select * from [kp_jsb] where [kp_item] like '%"&kp_item&"%' and [kp_zrrjs] like '%"&zrrjs&"%' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	Set tmpRs=xjweb.Exec(strSql, 1)
	do while not tmpRs.eof
		statkpjfz=statkpjfz + tmpRs("kp_uprice") * tmpRs("kp_mul")
		tmpRs.movenext
	loop
	tmpRs.close
	strSql="Select * from [ims_user] where mid(user_able,4,1)>0 and user_Group>0 and user_Group<4"
	Call xjweb.Exec("",-1)
	Set tmpRs=Server.CreateObject("ADODB.RECORDSET")
	tmpRs.open strSql,Conn,1,3
		PjCs=tmpRs.RecordCount
	tmpRs.close
	statkpjfz=Round(statkpjfz/PjCs,2)
End Function

Function statkpjcs(kp_item,zrrjs,i)
	Dim tmpRs
	statkpjcs=0
	strSql="select * from [kp_jsb] where [kp_item] like '%"&kp_item&"%' and [kp_zrrjs] like '%"&zrrjs&"%' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	Call xjweb.Exec("",-1)
	Set tmpRs=Server.CreateObject("ADODB.RECORDSET")
	tmpRs.open strSql,Conn,1,3
		statkpjcs=tmpRs.RecordCount
	tmpRs.close
	set tmprs=nothing
End Function

Function statkpfz(kp_item, i)
	statkpfz=0
	Dim tmpRs
	Select Case i
		Case 0		'对组员进行统计
			strSql="select * from [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case -1		'对主任进行统计
			strSql="select * from [kp_jsb] where [kp_item]='"&kp_item&"'  and [kp_kpr]<>" & struser & " and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'对组长进行统计
			strSql="select * from [kp_jsb] where [kp_group]="&i&"  and [kp_kpr]<>" & struser & " and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End Select

	Set tmpRs=xjweb.Exec(strSql, 1)

	do while not tmpRs.eof
		statkpfz=statkpfz + tmpRs("kp_uprice") * tmpRs("kp_mul")
		tmpRs.movenext
	loop
	statkpfz=round(statkpfz,2)
	tmpRs.close
	set tmprs=nothing
End Function

Function statkpcs(kp_item, kp_zrrjs, i)
	Dim TmpRs
	statkpcs=0
	If kp_zrrjs="" Then
		Select Case i
			Case 0		'对组员进行统计
				strSql=" [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'对主任进行统计
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'对组长进行统计
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 order by [kp_lsh]"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strsql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	else
		Select Case i
			Case 0		'对组员进行统计
				strSql=" [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'对主任进行统计
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'对组长进行统计
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and [kp_zrrjs]='"&kp_zrrjs&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strsql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	End If
End Function

Function diskpItem(arg1,arg2,arg3,arg4)
	icount=icount+1
	dim tmpcs, tmpkpf
	%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>&nbsp;</td>
  <td class=ltd><%=arg1%></td>
  <td class=ctd><%=arg2%></td>
  <td class=ctd><%=arg3%></td>
  <td class=ctd>分/项目(次数)</td>
  <%
					tmpcs=statkpcs(arg1, "", arg4)
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
				%>
  <td class=ctd><%=tmpcs%></td>
  <td class=ctd><%=tmpkpf%></td>
  <td class=ctd><%=kpf(icount-1)%></td>
</tr>
<%
End Function

Function diskpItemM(arg1,arg2,arg3,arg4,arg5)
	icount=icount+1
	dim tmpcs, tmpkpf, temparg
	temparg=arg1
	tmpcs=0
	tmpkpf=0
	If Instr(arg1,"原因产生返修")>0 Then temparg="原因产生返修"
	If Instr(arg1,"原因产生报废")>0 Then temparg="原因产生报废"
	strSql=""
	If arg4="" Then
		strSql="select * from [kp_jsb] where kp_item like '%"&temparg&"%' and kp_zrrjs='"&arg5&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	else
		strSql="select * from [kp_jsb] where kp_group="&arg4&" and kp_item like '%"&temparg&"%' and kp_zrrjs='"&arg5&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End If
	Set Rs = Server.CreateObject("ADODB.RECORDSET")
	Call xjweb.Exec("",-1)
	Rs.open strSql,Conn,3,3
  	tmpcs=rs.recordcount
	%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>&nbsp;</td>
  <td class=ltd><%=arg1%></td>
  <td class=ctd><%=arg2%>&nbsp;</td>
  <td class=ctd><%=arg3%></td>
  <td class=ctd>分/次数</td>
  <%
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If arg2<>"" Then
						If kpf(icount-1)<arg2*-1 Then
							 kpf(icount-1)=arg2*-1
						End If
					End If
					'组长返修、报废无上限
				%>
  <td class=ctd><%=tmpcs%></td>
  <td class=ctd><%=tmpkpf%></td>
  <td class=ctd><%=kpf(icount-1)%></td>
</tr>
<%
	Rs.Close
End Function
%>
