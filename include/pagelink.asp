<%
	'-------------------------------------定义各个链接项--------------------------------------------
	Function PageLink(pName)
		PageLink=""
		pName=LCase(pName)
		Select Case pName
			Case "index"		Rem 首页链接
				PageLink=PageLink & " <a href=mtask.asp>设计任务</a> |"
				PageLink=PageLink & " <a href=atask.asp>调试任务</a> |"
				PageLink=PageLink & " <a href=ftask.asp>零星任务</a> |"
				PageLink=PageLink & " <a href=mtest.asp>模具调试</a> |"
				PageLink=PageLink & " <a href=mtstat.asp>分值统计</a> |"
				PageLink=PageLink & " <a href=mquality.asp>模具质量</a> |"
				PageLink=PageLink & " <a href=docbak.asp>图档备份</a> |"
				PageLink=PageLink & " <a href=inform.asp>系统通知</a> |"
				PageLink=PageLink & " <a href=notebook.asp>系统留言</a> |"
				PageLink=PageLink & " <a href=uctrl.asp>用户设置</a> |"
				PageLink=PageLink & " <a href=/bbs>挤模论坛</a> "
				PageLink=""
			Case "mtask"		Rem 设计任务模块
				If ChkAble(3) Then PageLink=PageLink & " <a href=mtask_add.asp>添加设计任务书</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=jsdb_add.asp>添加代表任务书</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=mtask_change.asp>更改设计任务书</a> |"
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=mtask_assign.asp>分配任务书</a> |"
				If ChkAble(4) Then PageLink=PageLink & " <a href=mtask_zzchange.asp>更改责任人</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=mtask_delete.asp>删除任务书</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtask_display.asp>查看任务书</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=my_task.asp>我的任务</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtask_list.asp>任务流程</a>|"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=jsdb_list.asp>技术代表</a>"
			Case "atask"		Rem 调试任务模块
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=atask_changexs.asp>修改调试任务分值系数</a> |"
				If ChkAble("3,4,6") Then PageLink=PageLink & " <a href=atask_assign.asp>分配调试任务</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=atask_zzchange.asp>更改调试责任人</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=atask_display.asp>查看调试任务</a> || "
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=InfoFix_add.asp>齐套信息整理任务</a> |"
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=InfoFix_zzchange.asp>更改责任人</a> || "
				If ChkAble(-1) Then PageLink=PageLink & " <a href=atask_list.asp>调试任务列表</a>"
			Case "ftask"		Rem 零星任务模块
				If ChkAble("3,10") Then PageLink=PageLink & "<a href=ftask_add.asp>添加零星任务</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=ftask_list.asp>零星任务列表</a>"
			Case "mtest"		Rem 模具调试模块
				If ChkAble(6) Then PageLink=PageLink & "<a href=mtest_add.asp>添加调试信息</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtest_display.asp>查看调试信息</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtest_list.asp>调试信息总表</a> |"
				If ChkAble("1,2,3") Then PageLink=PageLink & " <a href=mtest_kp.asp>调试考评列表</a>"
'			Case "mquality"	Rem 模具质量
'				If ChkAble(0) Then PageLink=PageLink & "<a href=mquality_add.asp>添加质量信息</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_change.asp>更改质量信息</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_del.asp>删除质量信息</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_display.asp>查看质量信息</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_list.asp>质量信息总表</a>"
			Case "mtstat"		Rem 分值统计
				If ChkAble(0) Then PageLink=PageLink & "<a href=mtstat_display.asp>查看任务分值</a> |"
				If ChkAble(0) Then PageLink=PageLink & " <a href=mtstat_ygkpdis.asp>查看考评分值</a> |"
				If ChkAble("2,3") Then PageLink=PageLink & " <a href=mtstat_ygxslist.asp>查看员工系数</a> |"
				If ChkAble("2,3") Then PageLink=PageLink & " <a href=team_task.asp>任务定额</a>"
				
			Case "docbak"	Rem 图档备份
				If ChkAble(7) Then PageLink=PageLink & "<a href=docbak_add.asp>添加存档信息</a> |"
				If ChkAble(7) Then PageLink=PageLink & " <a href=docbak_change.asp>更改存档信息</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=docbak_search.asp>存档信息查询</a>"
			Case "tech"		Rem 问题分析
				If ChkAble(7) Then PageLink=PageLink & "<a href=tech_add.asp>添加问题分析</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=tech_display.asp>查看问题分析</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=tech_list.asp>问题分析列表</a> <p>"
				If ChkAble(11) Then PageLink=PageLink & "<a href=quality_add.asp>添加外部质量信息</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=quality_dis.asp>查看外部质量信息</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=quality_list.asp>外部质量信息列表</a> | "
				If ChkAble(11) Then PageLink=PageLink & "<a href=rectify_add.asp>添加纠正/预防措施</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=rectify_dis.asp>查看纠正/预防措施</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=rectify_list.asp>纠正/预防措施列表</a> </p>"
			Case "ygkp"		Rem 质量与考评
				If ChkAble("1,2,3,4,11,12") Then PageLink=PageLink & "<a href=ygkp_add.asp>添加考评</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=ygkp_list.asp>考评列表</a>"
			Case "inform"		Rem 通知
				If ChkAble("1,2,3") Then PageLink=PageLink & "<a href=inform_add.asp>发布通知</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=inform_dis.asp>查看通知</a>"
			Case "notebook"	 Rem 系统留言
				If ChkAble(0) Then PageLink=PageLink & "<a href=notebook_add.asp>撰写留言</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=notebook.asp>查看留言</a>"
			Case "uctrl"		Rem 用户控制面板
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_sendmsg.asp>发送短信</a> | "
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_dismsg.asp?box=incept>收件箱</a> | "
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_dismsg.asp?box=send>发件箱</a> | "
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_changeinf.asp>更改信息</a>"
			Case "gtask"	 Rem 工艺流程
				If ChkAble("3,4") Then PageLink=PageLink & "<a href=gtask_assign.asp>工艺流程</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtask_list.asp>设计流程</a>"
			Case Else			Rem 其他
				PageLink=pName & "(暂时没有可见链接)"
		End Select
	End Function
%>