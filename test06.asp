<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="分值统计 → Test06"					'页面的名称位置( 任务书管理 → 添加任务书)
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
	Dim kpf(30), kpif(5), ics(5), kpzf
	kpzf=0
	for i=0 to 29
		kpf(i)=0
	next
	for i=0 to 4
		kpif(i)=0
	next
	for i=0 to 4
		ics(i)=0
	next

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchMantime()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call ygkpstatDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub
	
Function SearchMantime()
%>
	<table cellpadding=2 cellspacing=0>
		<form action=<%=request.servervariables("script_name")%> method=get>
		<tr>
			<td>
			请选择:
			<select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
				<%for i = year(now) - 3 to year(now)%>
					<option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>年
			<select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
				<%for i = 1 to 12%>
					<option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>月&nbsp;&nbsp;

			<select name="searchuser" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
				<option value=""></option>
				<%If chkable("1,2,3,4") Then%>
					<%for i = 0 to ubound(c_allstat)%>
						<option value="<%=c_allstat(i)%>" <%If struser = c_allstat(i) Then%>selected<%end If%>><%=c_allstat(i)%></option>
					<%next%>
				<%Else%>
					<option value="<%=session("userName")%>"><%=session("userName")%></option>
				<%end If%>
			</select>&nbsp;
			<input type="submit" value=" 选 择 ">
			</td>
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

	Dim iTotalFz			'定义总分的变量

	If ChkJs(tmpAble)=5 Or ChkJs(tmpAble)=6 Then		'判断是不是组员或调试员
		'如果是组员或调试员的话在这里统计任务分值
		'统计分值
		'1--任务分值
		strSql="select a.*, b.* ,a.lsh as lsh, a.rwlr as rwlr from [mantime] a, [mtask] b where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 and a.lsh=b.lsh order by jssj desc, a.lsh desc"
		Set Rs=xjweb.Exec(strSql, 1)
		Dim itmpLsh, bJc		'奖惩临时变量
		itmpLsh="" : bJc=True
		If Not(Rs.eof or Rs.bof) Then 
			Do While Not Rs.eof
				If itmpLsh<>Rs("lsh") Then
					itmpLsh=Rs("lsh")
					bJc=True
				Else
					bJc=False
				End If
				irwzf=irwzf+Round(Rs("fz"),1)
				If bJc Then iaddfz=iaddfz+Rs("jc")
				Rs.movenext
			Loop
		End If
		Rs.close

		'2---零星任务分值
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by jssj desc"
		Set Rs=xjweb.Exec(strSql, 1)
		If Not(Rs.eof or Rs.bof) Then 
			Do While Not Rs.eof
				ilxrwzf=ilxrwzf+Rs("zf")
				Rs.movenext
			Loop
		End If
		Rs.close

		'3---统计总分
		
		If Fix(ilxrwzf + irwzf + iaddfz)<(ilxrwzf + irwzf + iaddfz) Then
			iTotalFz=Fix(ilxrwzf + irwzf + iaddfz) + 1
		Else
			iTotalFz=Fix(ilxrwzf + irwzf + iaddfz)
		End If

	End If		'判断是不是组员和调试员结束

	
	
	icount=1

	Call TbTopic(struser & " " & formatdatetime(dtstart,1) & " 至 " & formatdatetime(dtend,1) & " 考评统计")
		%>
		<table width="90%" cellpadding=2 cellspacing=0 class="xtable">
	<%
	Select Case ChkJs(tmpAble)
		Case 3	'主任
			%>
			<%Call diskpItem("设计结束文件未存盘",5.0,2.5, 0)%>
			<%Call diskpItem("调试结束文件未存盘",5.0,2.5, 0)%>
			<%Call diskpItem("技术文件签署、更改不完善",2.0,1.0, -1)%>
			<%Call diskpItem("文件更改未存档",2.0,2.0, 0)%>
			<%Call diskpItem("项目开发完成情况",6.0,2.0, -1)%>
			<%Call diskpItem("新品开发月、周未及时上报",2.0,1.0, -1)%>
			<%Call diskpItem("新品开发实施计划维护情况差",2.0,1.0, -1)%>
			<%Call diskpItem("设计规范维护",4.0,2.0, 0)%>
			<%Call diskpItem("不同批次相似型材类似问题重复",6.0,2.0, -1)%>
			<%Call diskpItem("厂内调试信息整理",4.0,1.0, -1)%>
			<%Call diskpItem("模具调试未合格数",4.0,2.0, 0)%>
			<%Call diskpItem("客户技术质理投诉与抱怨",4.0,2.0, 0)%>
			<%Call diskpItemM("设计原因产生返修",8.0,0.4, -1, 2)%>
			<%Call diskpItemM("设计原因产生报废",6.0,0.4, -1, 2)%>
			<%Call diskpItem("标准结构物料未充分利用",4.0,2.0, -1)%>
			<%Call diskpItem("成型方向腔数结构及速度和任务书不符",6.0,3.0, -1)%>
			<%Call diskpItem("接口件加热板胶块等和任务书不符",4.0,2.0, -1)%>
			<%Call diskpItem("同批次产品相同部位模具设计不一致",4.0,2.0, -1)%>
			<%Call diskpItem("设计准时完成率",12.0,2.0, 0)%>
			<%Call diskpItem("处理问题不及时",1.0,1.0, -1)%>
			<%Call diskpItem("部门协作",2.0,2.0, 0)%>
			<%Call diskpItem("工作日报周报月报",1.0,1.0, 0)%>
			<%Call diskpItem("月分析会整改项目",1.0,1.0, 0)%>
			<%Call diskpItem("质量工作计划完成",1.0,1.0, 0)%>
			<%Call diskpItem("纠正预防措施的执行",1.0,1.0, 0)%>
			<%Call diskpItem("使用非有效文件",1.0,1.0, 0)%>
			<%Call diskpItem("现场管理被厂或公司通报",1.0,1.0, 0)%>
			<%Call diskpItem("员工培训含技术交流",1.0,1.0, 0)%>
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			</tr>
			<%
		Case 4	'组长
			%>
			<%Call diskpItem("文件未按指定路径存盘",5.0,2.5, 0)%>
			<%Call diskpItem("技术文件签署、更改不完善",2.0,1.0, tmpGroup)%>
			<%Call diskpItem("新品开发月计划任务未完成",6.0,2.0, tmpGroup)%>
			<%Call diskpItem("模具设计因评审不充分返工",6.0,3.0, 0)%>
			<%Call diskpItem("标准结构物料未充分利用",4.0,2.0, tmpGroup)%>
			<%Call diskpItem("不同批次相似型材类似问题重复",6.0,2.0, 0)%>
			<%Call diskpItem("厂内调试信息整理",4.0,2.0, 0)%>
			<%Call diskpItemM("设计原因产生返修",6.0,0.6, tmpGroup, 2)%>
			<%Call diskpItemM("设计原因产生报废",8.0,0.6, tmpGroup, 2)%>
			<%Call diskpItemM("线切割成型1:1图",4.0,0.6, tmpGroup, 2)%>
			<%Call diskpItem("结构设计评审组织不及时",4.0,1.0, 0)%>
			<%Call diskpItem("成型方向腔数结构及速度和任务书不符",8.0,2.0, 0)%>
			<%Call diskpItem("接口件加热板胶块等和任务书不符",4.0,2.0, 0)%>
			<%Call diskpItem("同批次产品相同部位模具设计不一致",6.0,3.0, 0)%>
			<%Call diskpItem("设计准时完成率",12.0,4.0, 0)%>
			<%Call diskpItem("处理问题不及时",2.0,1.0, 0)%>
			<%Call diskpItem("工作周报",2.0,1.0, 0)%>
			<%Call diskpItem("小组整改",1.0,1.0, 0)%>
			<%Call diskpItem("分配质量工作计划完成情况",2.0,1.0, 0)%>
			<%Call diskpItem("分配纠正预防措施的执行",2.0,1.0, 0)%>
			<%Call diskpItem("员工培训含技术交流",1.0,1.0, 0)%>
			<%Call diskpItem("现场管理凌乱",1.0,1.0, 0)%>
			<%Call diskpItem("组员投诉",2.0,1.0, 0)%>
			<%Call diskpItem("分配零星任务完成情况",2.0,1.0, 0)%>
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			</tr>
			<%
		Case 6	'调试员
					If iTotalFz<300 Then
						kpf(0)=round((iTotalFz/300 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-300)*0.25)),1)
					End If
				%>
				<td class=ctd alt="<%="任务书:" & iTotalFz & "分"%>"><%=kpf(0)%></td>
				<%
					ics(0)=statkpcs("延迟", 0)
					ics(1)=statkpcs("提前", 0)					
					kpif(0)=statkpfz("延迟", 0)
					kpif(1)=statkpfz("提前", 0)
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("5次以上修理未评审", 0)
					ics(1)=statkpcs("修理方案下发不及时", 0)
					ics(2)=statkpcs("修理情报录入不及时", 0)
					ics(3)=statkpcs("修理图纸签署、更改不完善", 0)					
					kpif(0)=statkpfz("5次以上修理未评审", 0)
					kpif(1)=statkpfz("修理方案下发不及时", 0)
					kpif(2)=statkpfz("修理情报录入不及时", 0)
					kpif(3)=statkpfz("修理图纸签署、更改不完善", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(2)<-8 Then kpf(2)=-8
					ics(0)=statkpcs("修理图未及时存档", 0)
					ics(1)=statkpcs("修理方案原因产生返修", 0)
					ics(2)=statkpcs("修理方案原因产生报废", 0)
					kpif(0)=statkpfz("修理图未及时存档", 0)
					kpif(1)=statkpfz("修理方案原因产生返修", 0)
					kpif(2)=statkpfz("修理方案原因产生报废", 0)
					kpf(3)=kpif(0)+kpif(1)+kpif(2)
					If kpf(3)<-20 Then kpf(3)=-20
					ics(0)=statkpcs("第二次样品合格", 0)
					ics(1)=statkpcs("第三次样品合格", 0)
					ics(2)=statkpcs("模具调试未合格数", 0)
					kpif(0)=statkpfz("第二次样品合格", 0)
					kpif(1)=statkpfz("第三次样品合格", 0)
					kpif(2)=statkpfz("模具调试未合格数", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-6 Then kpf(4)=-6
					ics(0)=statkpcs("上班做与工作无关", 0)
					ics(1)=statkpcs("值班离岗,闲聊", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", 0)
					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("值班离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
					ics(0)=statkpcs("不服从分配", 0)
					ics(1)=statkpcs("消极怠工", 0)
					kpif(0)=statkpfz("不服从分配", 0)
					kpif(1)=statkpfz("消极怠工", 0)
					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				%>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+50
				%>
				<td class=ctd><%=kpzf%></td>
			<%
		Case 8		'技术指导负责人
			%>
			<%Call diskpItem("设计规范及标准升级维护",15.0,5.0, 0)%>
			<%Call diskpItem("模具设计因评审不充分返工",9.0,3.0, -1)%>
			<%Call diskpItem("标准结构物料未充分利用",16.0,4.0, -1)%>
			<%Call diskpItem("同批次产品相同部位模具设计不一致",12.0,4.0, -1)%>
			<%Call diskpItem("不同批次相似型材类似问题重复",16.0,4.0, -1)%>
			<%Call diskpItemM("设计原因产生返修",8.0,2.0, -1, 2)%>
			<%Call diskpItemM("设计原因产生报废",12.0,4.0, -1, 2)%>
			<%Call diskpItem("成型方向腔数结构及速度和任务书不符",8.0,2.0, -1)%>
			<%Call diskpItem("接口件加热板胶块等和任务书不符",4.0,2.0, -1)%>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			<%
		Case 5	'组员
					If iTotalFz<300 Then
						kpf(0)=round((iTotalFz/300 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-300)*0.25)),1)
					End If
				%>
				<td class=ctd alt="<%="任务书:" & iTotalFz & "分"%>"><%=kpf(0)%></td>
				<%
					ics(0)=statkpcs("延迟", 0)
					ics(1)=statkpcs("提前", 0)
					kpif(0)=statkpfz("延迟", 0)
					kpif(1)=statkpfz("提前", 0)
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("新品开发月计划任务未完成", 0)
					ics(1)=statkpcs("新品开发月、周未及时上报", 0)
					ics(2)=statkpcs("新品开发实施计划维护情况差", 0)
					kpif(0)=statkpfz("新品开发月计划任务未完成", 0)
					kpif(1)=statkpfz("新品开发月、周未及时上报", 0)
					kpif(2)=statkpfz("新品开发实施计划维护情况差", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2)
					If kpf(2)<-8 Then kpf(2)=-8
					ics(0)=statkpcs("技术文件签署、更改不完善", 0)
					kpif(0)=statkpfz("技术文件签署、更改不完善", 0)
					kpf(3)=kpif(0)
					If kpf(3)<-4 Then kpf(3)=-4
					ics(0)=statkpcs("标准结构物料未充分利用", 0)
					ics(1)=statkpcs("线切割成型1:1图", 0)
					ics(2)=statkpcs("设计原因产生返修", 0)
					ics(3)=statkpcs("设计原因产生报废", 0)
					kpif(0)=statkpfz("标准结构物料未充分利用", 0)
					kpif(1)=statkpfz("线切割成型1:1图", 0)
					kpif(2)=statkpfz("设计原因产生返修", 0)
					kpif(3)=statkpfz("设计原因产生报废", 0)
					kpf(4)=kpif(0)+kpif(1)+kpif(2)+kpif(3)
					If kpf(4)<-22 Then kpf(4)=-22
					ics(0)=statkpcs("第一次样品合格", 0)
					ics(1)=statkpcs("第二次样品合格", 0)
					kpif(0)=statkpfz("第一次样品合格", 0)
					kpif(1)=statkpfz("第二次样品合格", 0)
					kpf(5)=kpif(0) + kpif(1)
					ics(0)=statkpcs("上班做与工作无关", 0)
					ics(1)=statkpcs("值班离岗,闲聊", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", 0)
					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("值班离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
					ics(0)=statkpcs("不服从分配", 0)
					ics(1)=statkpcs("消极怠工", 0)
					kpif(0)=statkpfz("不服从分配", 0)
					kpif(1)=statkpfz("消极怠工", 0)
					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				%>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+50
				%>
				<td class=ctd><%=kpzf%></td>			
			<%
		Case 9			'项目管理负责人
			%>
			<%Call diskpItem("已结束项目未存盘",10.0,2.5, 0)%>
			<%Call diskpItem("项目文件签署不标准",15.0,3.0, 0)%>
			<%Call diskpItem("项目开发月报及时报送",15.0,3.0, 0)%>
			<%Call diskpItem("项目月报内容准时完成",15.0,3.0, 0)%>
			<%Call diskpItem("项目开发计划调整情况",10.0,5.0, 0)%>
			<%Call diskpItem("项目评审内容完成情况",20.0,4.0, 0)%>
			<%Call diskpItem("项目整改问题未完成",15.0,5.0, 0)%>
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			</tr>
			<%
		Case 10		'副主任
			%>
			<%Call diskpItem("设计结束文件未存盘",5.0,2.5, -1)%>
			<%Call diskpItem("调试结束文件未存盘",5.0,2.5, -1)%>
			<%Call diskpItem("技术文件签署、更改不完善",2.0,1.0, -1)%>
			<%Call diskpItem("文件更改未存档",1.0,1.0, -1)%>
			<%Call diskpItem("项目开发完成情况",3.0,1.0, -1)%>
			<%Call diskpItem("新品开发月、周未及时上报",2.0,1.0, -1)%>
			<%Call diskpItem("新品开发实施计划维护情况差",2.0,1.0, -1)%>
			<%Call diskpItem("设计规范维护",6.0,2.0, 0)%>
			<%Call diskpItem("不同批次相似型材类似问题重复",6.0,2.0, -1)%>
			<%Call diskpItem("厂内调试信息整理",6.0,2.0, -1)%>
			<%Call diskpItem("模具调试未合格数",6.0,2.0, 0)%>
			<%Call diskpItem("客户技术质理投诉与抱怨",3.0,1.0, 0)%>
			<%Call diskpItemM("设计原因产生返修",8.0,1.0, -1,2)%>
			<%Call diskpItemM("设计原因产生报废",6.0,3.0, -1,2)%>
			<%Call diskpItem("标准结构物料未充分利用",4.0,2.0, -1)%>
			<%Call diskpItem("成型方向腔数结构及速度和任务书不符",6.0,2.0, -1)%>
			<%Call diskpItem("接口件加热板胶块等和任务书不符",4.0,2.0, -1)%>
			<%Call diskpItem("同批次产品相同部位模具设计不一致",6.0,2.0, -1)%>
			<%Call diskpItem("设计准时完成率",12.0,4.0, -1)%>
			<%Call diskpItem("处理问题不及时",1.0,1.0, -1)%>
			<%Call diskpItem("部门协作",1.0,1.0, 0)%>
			<%Call diskpItem("工作日报周报月报",1.0,1.0, 0)%>
			<%Call diskpItem("月分析会整改项目",1.0,1.0, 0)%>
			<%Call diskpItem("纠正预防措施的执行",1.0,1.0, 0)%>
			<%Call diskpItem("使用非有效文件",1.0,1.0, 0)%>
			<%Call diskpItem("现场管理被厂或公司通报",1.0,1.0, 0)%>
			<%Call diskpItem("员工培训含技术交流",1.0,1.0, 0)%>
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
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

Function statkpcs(kp_item, i)
	statkpcs=0
	Select Case i
		Case 0		'对组员进行统计
			strSql=" [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case -1		'对主任进行统计
			strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'对组长进行统计
			strSql=" [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End Select
	statkpcs=xjweb.rscount(strSql)
End Function


Function diskpItem(arg1,arg2,arg3, arg4)
	dim tmpcs, tmpkpf
					tmpcs=statkpcs(arg1, arg4)
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
End Function

Function diskpItemM(arg1,arg2,arg3, arg4, arg5)
	icount=icount+1
	dim tmpcs, tmpkpf
					If Instr(arg1,"设计原因产生返修")>0 or Instr(arg1,"线切割成型1:1图") Then arg3=2*arg3
					If Instr(arg1,"设计原因产生报废")>0  Then arg3=4*arg3
					tmpcs=Int(statkpcs(arg1, arg4)/arg5)
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
End Function
%>