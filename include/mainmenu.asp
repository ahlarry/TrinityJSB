<%
	'ʹ�ò˵���������������ļ�
	Call FileInc(0, "js/menu.js")
	Call FileInc(1, "styles/menu.css")
	'-------------------------------------��������˵���--------------------------------------------
	Dim mnu_mtask, mnu_atask, mnu_ftask, mnu_mtest, mnu_mtstat, mnu_mquality, mnu_docbak, mnu_inform, mnu_notebook
	Dim mnu_uctrl, mnu_styles, mnu_sitestat, mnu_tech, mnu_ygkp, mnu_index

	rem ��ҳ(mnu_index)
	mnu_index=""
	If chkable("-1") Then mnu_index=mnu_index & "<div class=menuitems><a href=./dbm>ģ����Ϣ����</a></div>"
	If chkable(0) Then mnu_index=mnu_index & "<div class=menuitems><a href=./Jsbϵͳ�޸���־.htm>ϵͳ��־</a></div>"

	rem �������˵�(mnu_mtask)
	mnu_mtask=""
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_add.asp>������������</a></div>"
'	if chkable(3) then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=Repair_add.asp>�������������</a></div>"		
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=jsdb_add.asp>��Ӵ���������</a></div>"
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_change.asp>�������������</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"	
	If chkable("3,4") Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_assign.asp>����������</a></div>"
	If chkable(4) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_zzchange.asp>����������</a></div>"
	If chkable(3) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_delete.asp>ɾ��������</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_display.asp>�鿴������</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=my_task.asp>�ҵ�����</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_list.asp>��������</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=jsdb_list.asp>��������</a></div>"
	If chkable(-1) Then mnu_mtask=mnu_mtask & "<div class=menuitems><a href=mtask_gj.asp>����ģ��</a></div>"

	rem ��������˵�(mnu_atask)
	mnu_atask=""
	if chkable("3,4") then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_changexs.asp>�޸ĵ��������ֵϵ��</a></div>"
	if chkable("3,4,6") then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_assign.asp>�����������</a></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_display.asp>�鿴��������</a></div>"
	if chkable(4) then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_zzchange.asp>����������</a></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable("3,4") then mnu_atask=mnu_atask & "<div class=menuitems><a href=InfoFix_add.asp>������Ϣ��������</a></div>"
	if chkable("3,4") then mnu_atask=mnu_atask & "<div class=menuitems><a href=InfoFix_zzchange.asp>����������</a></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_atask=mnu_atask & "<div class=menuitems><a href=atask_list.asp>���������б�</a></div>"

	rem ��������˵�(mnu_ftask)
	mnu_ftask=""
	if chkable("3,10") then mnu_ftask=mnu_ftask & "<div class=menuitems><a href=ftask_add.asp>�����������</a></div>"
	if chkable("3,10") then mnu_ftask=mnu_ftask & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_ftask=mnu_ftask & "<div class=menuitems><a href=ftask_list.asp>���������б�</a></div>"

	rem ������Ϣ�˵�(mnu_mtest)
	mnu_mtest=""
	if chkable(6) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_add.asp>��ӵ�����Ϣ</a></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_display.asp>�鿴������Ϣ</a></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_list.asp>������Ϣ�ܱ�</a></div>"
	if chkable(-1) then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=Repair_list.asp>������Ϣ�ܱ�</a></div>"
	If chkable("1,2,3") Then mnu_mtest=mnu_mtest & "<div class=menuitems><a href=mtest_kp.asp>���Կ����б�</a></div>"

	rem ��ֵͳ�Ʋ˵�(mnu_mtstat)
	mnu_mtstat=""
	if chkable(0) then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=mtstat_display.asp>�鿴�����ֵ</a></div>"
	if chkable(0) then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=mtstat_ygkpdis.asp>�鿴������ֵ</a></div>"
	if chkable("2,3") then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=mtstat_ygxslist.asp>�鿴Ա��ϵ��</a></div>"
	If chkable("2,3") Then mnu_mtstat=mnu_mtstat & "<div class=menuitems><a href=team_task.asp>���񶨶�</a></div>"


	rem ͼ������(mnu_docbak)
	mnu_docbak=""
	if chkable(7) then mnu_docbak=mnu_docbak & "<div class=menuitems><a href=docbak_add.asp>��Ӵ浵��Ϣ</a></div>"
	if chkable(7) then mnu_docbak=mnu_docbak & "<div class=menuitems><a href=docbak_change.asp>���Ĵ浵��Ϣ</a></div>"
	if chkable(7) then mnu_docbak=mnu_docbak & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_docbak=mnu_docbak & "<div class=menuitems><a href=docbak_search.asp>�浵��Ϣ��ѯ</a></div>"

	rem �������(mnu_tech)
	mnu_tech=""
	if chkable(7) then mnu_tech=mnu_tech & "<div class=menuitems><a href=tech_add.asp>����������</a></div>"
	if chkable(-1) then mnu_tech=mnu_tech & "<div class=menuitems><a href=tech_display.asp>�鿴�������</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=tech_list.asp>��������б�</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	If chkable(11) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=quality_add.asp>����ⲿ������Ϣ</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=quality_list.asp>�ⲿ������Ϣ�б�</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=quality_dis.asp>�鿴�ⲿ������Ϣ</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuskin2><table width=80><tr><td class=td_frame height=1></td></tr></table></div>"
	If chkable(11) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=rectify_add.asp>��Ӿ���/Ԥ����ʩ</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=rectify_list.asp>����/Ԥ����ʩ�б�</a></div>"
	If chkable(-1) Then mnu_tech=mnu_tech & "<div class=menuitems><a href=rectify_dis.asp>�鿴����/Ԥ����ʩ</a></div>"

	rem ϵͳ֪ͨ(mnu_fdmail)
	mnu_inform=""
	if chkable(1) then mnu_inform=mnu_inform & "<div class=menuitems><a href=MayVoteAdmin/Admin_Login.asp>����ͶƱ</a></div>"
	if chkable("1,2,3") then mnu_inform=mnu_inform & "<div class=menuitems><a href=inform_add.asp>����֪ͨ</a></div>"
	if chkable(-1) then mnu_inform=mnu_inform & "<div class=menuitems><a href=inform_dis.asp>�鿴֪ͨ</a></div>"

	rem ϵͳ����
	mnu_notebook=""
	if chkable(0) then mnu_notebook=mnu_notebook & "<div class=menuitems><a href=notebook_add.asp>׫д����</a></div>"
	if chkable(-1) then mnu_notebook=mnu_notebook & "<div class=menuitems><a href=notebook.asp>�鿴����</a></div>"

	rem �û��������
	mnu_uctrl=""
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_sendmsg.asp>���Ͷ���</a></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_dismsg.asp?box=incept>�ռ���  </a></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_dismsg.asp?box=send>������  </a></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuskin2><table width=60><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(0) then mnu_uctrl=mnu_uctrl & "<div class=menuitems><a href=uctrl_changeinf.asp>������Ϣ</a></div>"

	rem Ա������
	mnu_ygkp=""
	if chkable("1,2,3,4,11,12") then mnu_ygkp=mnu_ygkp & "<div class=menuitems><a href=ygkp_add.asp>��ӿ���</a></div>"
	if chkable("1,2,3,4,11,12") then mnu_ygkp=mnu_ygkp & "<div class=menuskin2><table width=60><tr><td class=td_frame height=1></td></tr></table></div>"
	if chkable(-1) then mnu_ygkp=mnu_ygkp & "<div class=menuitems><a href=ygkp_list.asp>�����б�</a></div>"

	function mainmenu()
		dim strmmenutd,pcode			'strmmenutd---���˵�������
		pcode=""
		pcode=vbcrlf & "<!--ҳ�����˵����뿪ʼ--!>" &_
			vbcrlf & "<div class=menuskin id=popmenu onmouseover=""clearhidemenu();highlightmenu(event,'on')"" onmouseout=""highlightmenu(event,'off');dynamichide(event)"" style=""Z-index:100""></div>" &_
			vbcrlf & "<table border=0 cellspacing=0 cellpadding=0><tr>"

		strmmenutd="<td height=20 class=mmenu onmouseover='this.className=""mmenuover"";' onmouseout='this.className=""mmenu"";'>"

		pcode = pcode & strmmenutd &"<a onmouseover=""showmenu(event,'"&mnu_index&"')""  href=""index.asp"">��ҳ</a></td>"

		rem �������˵�
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mtask&"')"" style=""cursor:hand"" href=""mtask.asp"">�������</a></td>"

		rem ��������˵�
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_atask&"')"" style=""cursor:hand"" href=""atask.asp"">��������</a></td>"

		rem ��������˵�
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_ftask&"')"" style=""cursor:hand"" href="" ftask.asp"">��������</a></td>"

		rem ģ�ߵ���
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mtest&"')"" style=""cursor:hand"" href=""mtest.asp"">ģ�ߵ���</a></td>"

		rem ��ֵͳ��
		if chkable(0) then pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mtstat&"')"" style=""cursor:hand"" href=""mtstat.asp"">��ֵͳ��</a></td>"

		rem ģ������
		'pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_mquality&"')"" style=""cursor:hand;"" href=""mquality.asp"">ģ������</a></td>"

		rem ͼ������
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_docbak&"')"" style=""cursor:hand"" href=""docbak.asp"">ͼ������</a></td>"

		rem �����б�
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_tech&"')"" style=""cursor:hand"" href=""tech.asp"">�������</a></td>"

		rem Ա������
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_ygkp&"')"" style=""cursor:hand"" href=""ygkp.asp"">�����뿼��</a></td>"

		rem ϵͳ֪ͨ
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_inform&"')"" style=""cursor:hand"" href=""inform.asp"">ϵͳ֪ͨ</a></td>"

		rem ϵͳ����
		pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_notebook&"')"" style=""cursor:hand"" href=""notebook.asp"">ϵͳ����</a></td>"

		rem �û��������
		if chkable(0) then pcode = pcode & strmmenutd &"<a onMouseOver=""showmenu(event,'"&mnu_uctrl&"')"" style=""cursor:hand"" href=""uctrl.asp"">�û�����</a></td>"

		rem ��ģ��̳
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""/bbs"">��ģ��̳</a></td>"

		pcode =pcode & "</tr></table>" &_
			vbcrlf & "<!--ҳ�����˵��������--!>" & vbcrlf
		'response.write pcode
		mainmenu=pcode
	end function

	function bottommenu()		'�ײ�����(�˵�)
		dim strmmenutd	,pCode		'strmmenutd---���˵�������
		pcode=""
		pcode=vbcrlf & "<!--�ײ����Ӵ��뿪ʼ--!>" &_
			vbcrlf & "<table border=0 cellspacing=0 cellpadding=0><tr>"

		strmmenutd="<td height=20 class=mmenu onmouseover='this.className=""mmenuover"";' onmouseout='this.className=""mmenu"";'>"

		rem ��������
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""aboutus.asp"">��������</a>"
		rem IP����
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""ipman.asp"">IP����</a>"
		rem վ��ͳ��
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""sitestat.asp"">վ��ͳ��</a>"
		rem ϵͳ����
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""admin_index.asp"">ϵͳ����</a>"
		rem ʹ�ð���
		pcode = pcode & strmmenutd &"<a onmouseover=""hidemenu();"" href=""syshelp.asp"">ʹ�ð���</a>"


		pcode =pcode & "</tr></table>" &_
			vbcrlf & "<!--�ײ����Ӵ������--!>" & vbcrlf
		bottommenu=pcode
	end function
%>