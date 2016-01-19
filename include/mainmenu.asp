<%
	'使用菜单需包含下列两个文件
	Call FileInc(0, "js/menu.js")
	Call FileInc(1, "styles/menu.css")
	'-------------------------------------定义各个菜单项--------------------------------------------
	Dim mnu_mtask, mnu_atask, mnu_ftask, mnu_mtest, mnu_mtstat, mnu_mquality, mnu_docbak, mnu_inform, mnu_notebook
	Dim mnu_uctrl, mnu_styles, mnu_sitestat, mnu_tech, mnu_ygkp, mnu_index

	rem 首页(mnu_index)
	mnu_index=""
	If chkable("-1") Then mnu_index=mnu_index & "<div class=menuitems><a href=./dbm>模具信息管理</a></div>"
	If chkable(0) Then mnu_index=mnu_index & "<div class=menuitems><a href=./Jsb系统修改日志.htm>系统日志</a></div>"

	rem 设计任务菜单(mnu_mtask)
	mnu_mtask=""
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_add.asp>添加设计任务书</a></div>"
'	if chkable(3) then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=Repair_add.asp>添改修理任务书</a></div>"		
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=jsdb_add.asp>添加代表任务书</a></div>"
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_change.asp>更改设计任务书</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"	
	If chkable("3,4") Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_assign.asp>分配任务书</a></div>"
	If chkable(4) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_zzchange.asp>更改责任人</a></div>"
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_delete.asp>删除任务书</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_display.asp>查看任务书</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=my_task.asp>我的任务</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_list.asp>任务流程</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=jsdb_list.asp>技术代表</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_gj.asp>共挤模具</a></div>"

	rem 调试任务菜单(mnu_atask)
	mnu_atask=""
	if chkable("3,4") then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_changexs.asp>修改调试任务分值系数</a></div>"
	if chkable("3,4,6") then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_assign.asp>分配调试任务</a></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_display.asp>查看调试任务</a></div>"
	if chkable(4) then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_zzchange.asp>更改责任人</a></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable("3,4") then mnu_atask=mnu_atask & "<div class=menuitems><a href=InfoFix_add.asp>齐套信息整理任务</a></div>"
	if chkable("3,4") then mnu_atask=mnu_atask & "<div class=menuitems><a href=InfoFix_zzchange.asp>更改责任人</a></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_list.asp>调试任务列表</a></div>"

	rem 零星任务菜单(mnu_ftask)
	mnu_ftask=""
	if chkable("3,10") then mnu_ftask=mnu_ftask & "<div class=menuitems><a href=ftask_add.asp>添加零星任务</a></div>"
	if chkable("3,10") then mnu_ftask=mnu_ftask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_ftask=mnu_ftask & "<div class=menuitems><a href=ftask_list.asp>零星任务列表</a></div>"

	rem 调试信息菜单(mnu_mtest)
	mnu_mtest=""
	if chkable(6) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_add.asp>添加调试信息</a></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_display.asp>查看调试信息</a></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_list.asp>调试信息总表</a></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=Repair_list.asp>修理信息总表</a></div>"
	If chkable("1,2,3") Then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_kp.asp>调试考评列表</a></div>"

	rem 分值统计菜单(mnu_mtstat)
	mnu_mtstat=""
	if chkable(0) then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=mtstat_display.asp>查看任务分值</a></div>"
	if chkable(0) then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=mtstat_ygkpdis.asp>查看考评分值</a></div>"
	if chkable("2,3") then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=mtstat_ygxslist.asp>查看员工系数</a></div>"
	If chkable("2,3") Then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=team_task.asp>任务定额</a></div>"


	rem 图档备份(mnu_docbak)
	mnu_docbak=""
	if chkable(7) then mnu_docbak=mnu_docbak & "<div class=menuitems><a href=docbak_add.asp>添加存档信息</a></div>"
	if chkable(7) then mnu_docbak=mnu_docbak & "<div class=menuitems><a href=docbak_change.asp>更改存档信息</a></div>"
	if chkable(7) then mnu_docbak=mnu_docbak & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_docbak=mnu_docbak & "<div class=menuitems><a href=docbak_search.asp>存档信息查询</a></div>"

	rem 问题分析(mnu_tech)
	mnu_tech=""
	if chkable(7) then mnu_tech=mnu_tech & "<div class=menuitems><a href=tech_add.asp>添加问题分析</a></div>"
	if chkable(-1) then mnu_tech=mnu_tech & "<div class=menuitems><a href=tech_display.asp>查看问题分析</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=tech_list.asp>问题分析列表</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	If chkable(11) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=quality_add.asp>添加外部质量信息</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=quality_list.asp>外部质量信息列表</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=quality_dis.asp>查看外部质量信息</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	If chkable(11) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=rectify_add.asp>添加纠正/预防措施</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=rectify_list.asp>纠正/预防措施列表</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=rectify_dis.asp>查看纠正/预防措施</a></div>"

	rem 系统通知(mnu_fdmail)
	mnu_inform=""
	if chkable(1) then mnu_inform=mnu_inform & "<div class=menuitems><a href=MayVoteAdmin/Admin_Login.asp>发布投票</a></div>"
	if chkable("1,2,3") then mnu_inform=mnu_inform & "<div class=menuitems><a href=inform_add.asp>发布通知</a></div>"
	if chkable(-1) then mnu_inform=mnu_inform & "<div class=menuitems><a href=inform_dis.asp>查看通知</a></div>"

	rem 系统留言
	mnu_notebook=""
	if chkable(0) then mnu_notebook=mnu_notebook & "<div class=menuitems><a href=notebook_add.asp>撰写留言</a></div>"
	if chkable(-1) then mnu_notebook=mnu_notebook & "<div class=menuitems><a href=notebook.asp>查看留言</a></div>"

	rem 用户控制面板
	mnu_uctrl=""
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_sendmsg.asp>发送短信</a></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_dismsg.asp?box=incept>收件箱  </a></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_dismsg.asp?box=send>发件箱  </a></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuskin2><table width=60><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_changeinf.asp>更改信息</a></div>"

	rem 员工考评
	mnu_ygkp=""
	if chkable("1,2,3,4,11,12") then mnu_ygkp=mnu_ygkp & "<div class=menuitems><a href=ygkp_add.asp>添加考评</a></div>"
	if chkable("1,2,3,4,11,12") then mnu_ygkp=mnu_ygkp & "<div class=menuskin2><table width=60><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_ygkp=mnu_ygkp & "<div class=menuitems><a href=ygkp_list.asp>考评列表</a></div>"

	function mainmenu()
		dim strmmenutd,pcode			'strmmenutd---主菜单表格变量
		pcode=""
		pcode=vbcrlf & "<!--页面主菜单代码开始--!>" &_
			vbcrlf & "<div class=menuskin id=popmenu onmouseover=""clearhidemenu();highlightmenu(event,'on')"" onmouseout=""highlightmenu(event,'off');dynamichide(event)"" style=""Z-index:100""></div>" &_
			vbcrlf & "<table border=0 cellspacing=0 cellpadding=0><tr>"

		strmmenutd="<td height=20 class=mmenu onmouseover='this.className=""mmenuover"";' onmouseout='this.className=""mmenu"";'>"

		pcode = pcode & strmmenutd &"<a onmouseover=""showmenu(event,'"&mnu_index&"')""  href=""index.asp"">首页</a></td>"

		rem 设计任务菜单
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mtask&"')"" style=""cursor:hand"" href=""mtask.asp"">设计任务</a></td>"

		rem 调试任务菜单
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_atask&"')"" style=""cursor:hand"" href=""atask.asp"">调试任务</a></td>"

		rem 零星任务菜单
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_ftask&"')"" style=""cursor:hand"" href="" ftask.asp"">零星任务</a></td>"

		rem 模具调试
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mtest&"')"" style=""cursor:hand"" href=""mtest.asp"">模具调试</a></td>"

		rem 分值统计
		if chkable(0) then pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mtstat&"')"" style=""cursor:hand"" href=""mtstat.asp"">分值统计</a></td>"

		rem 模具质量
		'pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mquality&"')"" style=""cursor:hand;"" href=""mquality.asp"">模具质量</a></td>"

		rem 图档备份
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_docbak&"')"" style=""cursor:hand"" href=""docbak.asp"">图档备份</a></td>"

		rem 问题列表
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_tech&"')"" style=""cursor:hand"" href=""tech.asp"">问题分析</a></td>"

		rem 员工考评
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_ygkp&"')"" style=""cursor:hand"" href=""ygkp.asp"">质量与考评</a></td>"

		rem 系统通知
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_inform&"')"" style=""cursor:hand"" href=""inform.asp"">系统通知</a></td>"

		rem 系统留言
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_notebook&"')"" style=""cursor:hand"" href=""notebook.asp"">系统留言</a></td>"

		rem 用户控制面板
		if chkable(0) then pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_uctrl&"')"" style=""cursor:hand"" href=""uctrl.asp"">用户操作</a></td>"

		rem 挤模论坛
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""/bbs"">挤模论坛</a></td>"

		pcode =pcode & "</tr></table>" &_
			vbcrlf & "<!--页面主菜单代码结束--!>" & vbcrlf
		'response.write pcode
		mainmenu=pcode
	end function

	function bottommenu()		'底部链接(菜单)
		dim strmmenutd	,pCode		'strmmenutd---主菜单表格变量
		pcode=""
		pcode=vbcrlf & "<!--底部链接代码开始--!>" &_
			vbcrlf & "<table border=0 cellspacing=0 cellpadding=0><tr>"

		strmmenutd="<td height=20 class=mmenu onmouseover='this.className=""mmenuover"";' onmouseout='this.className=""mmenu"";'>"

		rem 关于我们
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""aboutus.asp"">关于我们</a>"
		rem IP管理
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""ipman.asp"">IP管理</a>"
		rem 站内统计
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""sitestat.asp"">站内统计</a>"
		rem 系统管理
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""admin_index.asp"">系统管理</a>"
		rem 使用帮助
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""syshelp.asp"">使用帮助</a>"


		pcode =pcode & "</tr></table>" &_
			vbcrlf & "<!--底部链接代码结束--!>" & vbcrlf
		bottommenu=pcode
	end function
%>