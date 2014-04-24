<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="分值统计 → 查看员工系数"					'页面的名称位置( 分值统计 → 查看员工系数)
strPage="mtstat"
xjweb.header()
Call TopTable()

Dim iyear, imonth, dtstart, dtend, irwzf, iaddfz, zcount, icount, ilxrwzf, zrwwcl, zgroup
Dim zuser, zrwfz, zrwxs, zzlxs, zgkxs, zbmxs, zjbgz, zjxgz,zyfgz, zbeiz, ygxsRs ,m
zjbgz=0
zjxgz=0
zyfgz=0
zgroup=0
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
<THEAD>
  <tr>
    <th class=th width="8%">ID</th>
    <th class=th width="10%">人员名单</th>
    <th class=th width="10%">任务分值</th>
    <th class=th width="10%">定量考核</th>
    <th class=th width="10%">定性考核</th>
    <th class=th width="10%">综合考核</th>
    <th class=th width="10%">部考系数</th>
    <th class=th width="10%">基本工资</th>
    <th class=th width="10%">绩效工资</th>
    <th class=th width="10%">应发工资</th>
  </tr>
  <tr>
  	<td colspan="10" class=rtd>本月部门任务完成率=<%=zrwwcl%></td>
  </tr>
  </THEAD>
  <%
		Dim strColor
		strColor=-1
		If Request("bybmxs")="" Then zbmxs=1.0 Else zbmxs=Request("bybmxs") End if
		strSql="select * from [ims_user] where user_depart='技术部' and user_group<>0 and user_able<>'010000000000000' and Instr('AABBTB调试员',user_name)=0 order by user_group,user_able"
		Set ygxsRs=xjweb.Exec(strSql, 1)
		Do While Not ygxsRs.eof	or ygxsRs.Bof
		zuser=ygxsRs("user_name")
		If zgroup<>ygxsRs("user_group") Then
			strColor=-1*strColor
			zgroup=ygxsRs("user_group")
		End If
		zgroup=ygxsRs("user_group")
		zrwfz=0 : zrwxs=0 : zzlxs=0 : zgkxs=0 : zjbgz=0 : zyfgz=0 : zbeiz="" : irwzf=0 : ilxrwzf=0 : iaddfz=0
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
<TBODY>
  <tr <%If strColor=1 Then%>bgcolor="#D6D7EF"<%End If%>>
    <td class=ctd width="8%"><%=zcount%></td>
    <td class=ctd width="10%"><%=zuser%></td>
    <td class=ctd width="10%"><%=zrwfz%>&nbsp;</td>
    <td class=ctd width="10%"><%=zrwxs%>&nbsp;</td>
    <td class=ctd width="10%"><%=zzlxs%>&nbsp;</td>
    <td class=ctd width="10%"><%=zgkxs%>&nbsp;</td>
    <td class=ctd width="10%"><%=zbmxs%></td>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="10%">&nbsp;</td>
  </tr>
</TBODY>
  <%
		zcount = zcount + 1
		ygxsRs.movenext
		loop
		ygxsRs.close
%>
<TFOOT>
<TR>
<TD class=rtd colspan=10>
The End.
</TD>
</TR>
</TFOOT>
</table>
<%
End Function

Function YgxsStat()
	strSql="Select * from [ims_user] where [user_name]='"&zuser&"'"
	Set Rs=xjweb.Exec(strSql,1)
	Dim tmpCount, tmpGroup, tmpAble, ilxrwzf
	tmpCount=1
	tmpGroup=Rs("user_Group")
	tmpAble=Rs("user_Able")
	Rs.Close

	If InStr("5689",ChkJs(tmpAble))>0 Then		'判断是不是组员或调试员
		'1--任务分值
		strSql="select * from [mantime] where zrr='"&zuser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			irwzf=irwzf+Round(Rs("fz"),1)
			Rs.movenext
		Loop
		Rs.close
		'2---零星任务分值
		strSql="select * from [ftask] where zrr='"&zuser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
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
			zrwfz=Round((ilxrwzf + irwzf)/tmpCount,1)
		End If
	End If
	zjxgz=1800
	icount=1
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
<%Call diskpItemM("设计原因产生返修",8.0,0.4, "", "设计")%>
<%Call diskpItemM("设计原因产生报废",6.0,0.4, "", "设计")%>
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
<%	for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				zjxgz=2800
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=zzlxs
			'	zyfgz=zjxgz*zgkxs*zrwwcl*zbmxs+zjbgz

		Case 4	'组长
					If zrwfz<300 Then
						kpf(0)=round((zrwfz/300 * 50),1)
					Else
						kpf(0)=round((50+((zrwfz-300)/300*50*1.25)),1)
					End If
					ics(0)=statkpcs("提前", "", 0)
					ics(1)=statkpcs("延迟", "", 0)
					kpif(0)=statkpfz("提前", 0)
					kpif(1)=statkpfz("延迟", 0)
					kpf(1)=kpif(0) + kpif(1)
					if kpf(1)<-10 Then kpf(1)=-10
					if kpf(1)>10 Then kpf(1)=10
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
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>300 Then
					kpzf=kpzf+30
				else
					kpzf=kpzf+(zrwfz/300 * 30)
				End If
				zrwxs=round(kpzf/100,2)
				zzlxs=round((20+kpf(3))/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
				zjxgz=2000
		Case 5	'组员
					If zrwfz<300 Then
						kpf(0)=round((zrwfz/300 * 50),1)
					Else
						kpf(0)=round((50+((zrwfz-300)/300*50*1.25)),1)
					End If
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)
					kpif(0)=statkpfz("延迟", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("提前", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
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
					ics(0)=statkpcs("提出改进建议取得成效", "", 0)
					ics(1)=statkpcs("提出合理化建议并被采纳", "", 0)
					kpif(0)=statkpfz("提出改进建议取得成效", 0)
					kpif(1)=statkpfz("提出合理化建议并被采纳", 0)
					kpf(3)=kpif(0) + kpif(1)
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("不服从分配", "", 0)
					ics(2)=statkpcs("主动承担较难任务", "", 0)
					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("不服从分配", 0)
					kpif(2)=statkpfz("主动承担较难任务", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-20 Then kpf(4)=-20
					If kpf(4)>20 Then kpf(4)=20
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>300 Then
					kpzf=kpzf+30
				else
					kpzf=kpzf+(zrwfz/300 * 30)
				End If
				zrwxs=round(kpzf/100,2)
				zzlxs=round((20+kpf(3)+kpf(4))/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
		Case 6	'调试员
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)
					kpif(0)=statkpfz("延迟", 0)
					kpif(1)=statkpfz("提前", 0)
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
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
					ics(0)=statkpcs("调试超过额定最大次数", "", 0)
					ics(1)=statkpcs("调试少于额定最小次数", "", 0)
					ics(2)=statkpcs("模具调试未合格数", "", 0)
					kpif(0)=statkpfz("调试超过额定最大次数", 0)
					kpif(1)=statkpfz("调试少于额定最小次数", 0)
					kpif(2)=statkpfz("模具调试未合格数", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-6 Then kpf(4)=-6
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("值班离岗,闲聊", "", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", "", 0)
					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("值班离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
					ics(0)=statkpcs("不服从分配", "", 0)
					ics(1)=statkpcs("消极怠工", "", 0)
					kpif(0)=statkpfz("不服从分配", 0)
					kpif(1)=statkpfz("消极怠工", 0)
					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>300 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(zrwfz/300 * 50),1)
				End If
				If zrwfz<300 Then
					zrwxs=round((zrwfz/300 * 50)/100,2)
				Else
					zrwxs=round(((50+((zrwfz-300)/300*50*1.25))/100),2)
				End If
				zzlxs=round(kpzf/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
				zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 7			'图档管理员
			zjxgz=600
			zzlxs=1.00
			zgkxs=1.00
		'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz


		Case 8		'工艺员
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)
					kpif(0)=statkpfz("延迟", 0)
					kpif(1)=statkpfz("提前", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("新品开发月计划任务未完成", "", 0)
					ics(1)=statkpcs("新品开发月、周未及时上报", "", 0)
					ics(2)=statkpcs("新品开发实施计划维护情况差", "", 0)
					kpif(0)=statkpfz("新品开发月计划任务未完成", 0)
					kpif(1)=statkpfz("新品开发月、周未及时上报", 0)
					kpif(2)=statkpfz("新品开发实施计划维护情况差", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2)
					If kpf(2)<-5 Then kpf(2)=-5
					ics(0)=statkpcs("工艺文件更改不完善未造成质量问题", "", 0)
					kpif(0)=statkpfz("工艺文件更改不完善未造成质量问题", 0)
					ics(1)=statkpcs("所编制的作业指导书经论证后未及时修改", "", 0)
					kpif(1)=statkpfz("所编制的作业指导书经论证后未及时修改", 0)
					ics(2)=statkpcs("工艺文件签署不完整", "", 0)
					kpif(2)=statkpfz("工艺文件签署不完整", 0)
					ics(3)=statkpcs("因签署日期不实，导致质量争议", "", 0)
					kpif(3)=statkpfz("因签署日期不实，导致质量争议", 0)
					ics(4)=statkpcs("因工艺错误造成返修（漏、错工艺）", "", 0)
					kpif(4)=statkpfz("因工艺错误造成返修（漏、错工艺）", 0)
					ics(5)=statkpcs("因工艺错误造成报废", "", 0)
					kpif(5)=statkpfz("因工艺错误造成报废", 0)
					ics(6)=statkpcs("未按规范执行", "", 0)
					kpif(6)=statkpfz("未按规范执行", 0)
					ics(7)=statkpcs("因工艺造成车间订单不能下达", "", 0)
					kpif(7)=statkpfz("因工艺造成车间订单不能下达", 0)
					ics(8)=statkpcs("相同错误重复出现", "", 0)
					kpif(8)=statkpfz("相同错误重复出现", 0)
					ics(9)=statkpcs("设计结构、材料、热处理等明显错误未及时反映", "", 0)
					kpif(9)=statkpfz("设计结构、材料、热处理等明显错误未及时反映", 0)
					kpf(3)=kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)+kpif(5)+kpif(6)+kpif(7)+kpif(8)+kpif(9)
					If kpf(3)<-20 Then kpf(3)=-20
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
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("值班离岗,闲聊", "", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", "", 0)
					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("值班离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)
					kpf(5)=kpif(0) + kpif(1) + kpif(2)
					If kpf(5)<-5 Then kpf(5)=-5
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
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>400 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(zrwfz/400 * 50),1)
				End If
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=FormatNumber(zrwxs+zzlxs,2)
			'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 9		'编程员
					ics(0)=statkpcs("延迟", "", 0)
					ics(1)=statkpcs("提前", "", 0)
					kpif(0)=statkpfz("延迟", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("提前", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
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
					ics(0)=statkpcs("上班做与工作无关", "", 0)
					ics(1)=statkpcs("无正当理由离岗,闲聊", "", 0)
					ics(2)=statkpcs("桌面不洁,下班机器未关、门未锁", "", 0)
					kpif(0)=statkpfz("上班做与工作无关", 0)
					kpif(1)=statkpfz("无正当理由离岗,闲聊", 0)
					kpif(2)=statkpfz("桌面不洁,下班机器未关、门未锁", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-2 Then kpf(4)=-2
					ics(0)=statkpcs("不服从分配", "", 0)
					ics(1)=statkpcs("处理问题不及时", "", 0)
					ics(2)=statkpcs("消极怠工", "", 0)
					kpif(0)=statkpfz("不服从分配", 0)
					kpif(1)=statkpfz("处理问题不及时", 0)
					kpif(2)=statkpfz("消极怠工", 0)
					kpf(5)=kpif(0) + kpif(1) + kpif(2)
					If kpf(5)<-5 Then kpf(5)=-5
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>400 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(zrwfz/400 * 50),1)
				End If
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=FormatNumber(zrwxs+zzlxs,2)
			'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

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
<%Call diskpItemM("设计原因产生返修",8.0,0.4, "","设计")%>
<%Call diskpItemM("设计原因产生报废",6.0,0.4, "","设计")%>
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
<%	for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=FormatNumber(zrwxs+zzlxs,2)
			'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 13		'网络管理员
				zzlxs=""
				zgkxs=""
				zyfgz=""

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
	strSql="Select * from [ims_user] where mid(user_able,4,1)>0 and user_Group>0 and user_Group<>5"
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
			strSql="select * from [kp_jsb] where [kp_zrr]='"&zuser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case -1		'对主任进行统计
			strSql="select * from [kp_jsb] where [kp_item]='"&kp_item&"'  and [kp_kpr]<>" & zuser & " and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'对组长进行统计
			strSql="select * from [kp_jsb] where [kp_group]="&i&"  and [kp_kpr]<>" & zuser & " and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
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
	Dim tmpRs
	statkpcs=0
	If kp_zrrjs="" Then
		Select Case i
			Case 0		'对组员进行统计
				strSql=" [kp_jsb] where [kp_zrr]='"&zuser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'对主任进行统计
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'对组长进行统计
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 order by [kp_lsh]"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strSql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	else
				Select Case i
			Case 0		'对组员进行统计
				strSql=" [kp_jsb] where [kp_zrr]='"&zuser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'对主任进行统计
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'对组长进行统计
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and [kp_zrrjs]='"&kp_zrrjs&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strSql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	End If
End Function

Function diskpItem(arg1,arg2,arg3,arg4)
	icount=icount+1
	dim tmpcs, tmpkpf
	tmpcs=statkpcs(arg1, "", arg4)
	tmpkpf=tmpcs*arg3*-1
	kpf(icount-1)=tmpkpf
	If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
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
	tmpkpf=tmpcs*arg3*-1
	kpf(icount-1)=tmpkpf
	If arg2<>"" Then
		If kpf(icount-1)<arg2*-1 Then
			 kpf(icount-1)=arg2*-1
		End If
	End If
	'组长返修、报废无上限
	Rs.Close
End Function
%>
