<%
	'-------------------------------------�������������--------------------------------------------
	Function PageLink(pName)
		PageLink=""
		pName=LCase(pName)
		Select Case pName
			Case "index"		Rem ��ҳ����
				PageLink=PageLink & " <a href=mtask.asp>�������</a> |"
				PageLink=PageLink & " <a href=atask.asp>��������</a> |"
				PageLink=PageLink & " <a href=ftask.asp>��������</a> |"
				PageLink=PageLink & " <a href=mtest.asp>ģ�ߵ���</a> |"
				PageLink=PageLink & " <a href=mtstat.asp>��ֵͳ��</a> |"
				PageLink=PageLink & " <a href=mquality.asp>ģ������</a> |"
				PageLink=PageLink & " <a href=docbak.asp>ͼ������</a> |"
				PageLink=PageLink & " <a href=inform.asp>ϵͳ֪ͨ</a> |"
				PageLink=PageLink & " <a href=notebook.asp>ϵͳ����</a> |"
				PageLink=PageLink & " <a href=uctrl.asp>�û�����</a> |"
				PageLink=PageLink & " <a href=/bbs>��ģ��̳</a> "
				PageLink=""
			Case "mtask"		Rem �������ģ��
				If ChkAble(3) Then PageLink=PageLink & " <a href=mtask_add.asp>������������</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=jsdb_add.asp>��Ӵ���������</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=mtask_change.asp>�������������</a> |"
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=mtask_assign.asp>����������</a> |"
				If ChkAble(4) Then PageLink=PageLink & " <a href=mtask_zzchange.asp>����������</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=mtask_delete.asp>ɾ��������</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtask_display.asp>�鿴������</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=my_task.asp>�ҵ�����</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtask_list.asp>��������</a>|"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=jsdb_list.asp>��������</a>"
			Case "atask"		Rem ��������ģ��
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=atask_changexs.asp>�޸ĵ��������ֵϵ��</a> |"
				If ChkAble("3,4,6") Then PageLink=PageLink & " <a href=atask_assign.asp>�����������</a> |"
				If ChkAble(3) Then PageLink=PageLink & " <a href=atask_zzchange.asp>���ĵ���������</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=atask_display.asp>�鿴��������</a> || "
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=InfoFix_add.asp>������Ϣ��������</a> |"
				If ChkAble("3,4") Then PageLink=PageLink & " <a href=InfoFix_zzchange.asp>����������</a> || "
				If ChkAble(-1) Then PageLink=PageLink & " <a href=atask_list.asp>���������б�</a>"
			Case "ftask"		Rem ��������ģ��
				If ChkAble("3,10") Then PageLink=PageLink & "<a href=ftask_add.asp>�����������</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=ftask_list.asp>���������б�</a>"
			Case "mtest"		Rem ģ�ߵ���ģ��
				If ChkAble(6) Then PageLink=PageLink & "<a href=mtest_add.asp>��ӵ�����Ϣ</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtest_display.asp>�鿴������Ϣ</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtest_list.asp>������Ϣ�ܱ�</a> |"
				If ChkAble("1,2,3") Then PageLink=PageLink & " <a href=mtest_kp.asp>���Կ����б�</a>"
'			Case "mquality"	Rem ģ������
'				If ChkAble(0) Then PageLink=PageLink & "<a href=mquality_add.asp>���������Ϣ</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_change.asp>����������Ϣ</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_del.asp>ɾ��������Ϣ</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_display.asp>�鿴������Ϣ</a> |"
'				If ChkAble(0) Then PageLink=PageLink & " <a href=mquality_list.asp>������Ϣ�ܱ�</a>"
			Case "mtstat"		Rem ��ֵͳ��
				If ChkAble(0) Then PageLink=PageLink & "<a href=mtstat_display.asp>�鿴�����ֵ</a> |"
				If ChkAble(0) Then PageLink=PageLink & " <a href=mtstat_ygkpdis.asp>�鿴������ֵ</a> |"
				If ChkAble("2,3") Then PageLink=PageLink & " <a href=mtstat_ygxslist.asp>�鿴Ա��ϵ��</a> |"
				If ChkAble("2,3") Then PageLink=PageLink & " <a href=team_task.asp>���񶨶�</a>"
				
			Case "docbak"	Rem ͼ������
				If ChkAble(7) Then PageLink=PageLink & "<a href=docbak_add.asp>��Ӵ浵��Ϣ</a> |"
				If ChkAble(7) Then PageLink=PageLink & " <a href=docbak_change.asp>���Ĵ浵��Ϣ</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=docbak_search.asp>�浵��Ϣ��ѯ</a>"
			Case "tech"		Rem �������
				If ChkAble(7) Then PageLink=PageLink & "<a href=tech_add.asp>����������</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=tech_display.asp>�鿴�������</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=tech_list.asp>��������б�</a> <p>"
				If ChkAble(11) Then PageLink=PageLink & "<a href=quality_add.asp>����ⲿ������Ϣ</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=quality_dis.asp>�鿴�ⲿ������Ϣ</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=quality_list.asp>�ⲿ������Ϣ�б�</a> | "
				If ChkAble(11) Then PageLink=PageLink & "<a href=rectify_add.asp>��Ӿ���/Ԥ����ʩ</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=rectify_dis.asp>�鿴����/Ԥ����ʩ</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=rectify_list.asp>����/Ԥ����ʩ�б�</a> </p>"
			Case "ygkp"		Rem �����뿼��
				If ChkAble("1,2,3,4,11,12") Then PageLink=PageLink & "<a href=ygkp_add.asp>��ӿ���</a> | "
				If ChkAble(-1) Then PageLink=PageLink & "<a href=ygkp_list.asp>�����б�</a>"
			Case "inform"		Rem ֪ͨ
				If ChkAble("1,2,3") Then PageLink=PageLink & "<a href=inform_add.asp>����֪ͨ</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=inform_dis.asp>�鿴֪ͨ</a>"
			Case "notebook"	 Rem ϵͳ����
				If ChkAble(0) Then PageLink=PageLink & "<a href=notebook_add.asp>׫д����</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=notebook.asp>�鿴����</a>"
			Case "uctrl"		Rem �û��������
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_sendmsg.asp>���Ͷ���</a> | "
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_dismsg.asp?box=incept>�ռ���</a> | "
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_dismsg.asp?box=send>������</a> | "
				If ChkAble(0) Then PageLink=PageLink & "<a href=uctrl_changeinf.asp>������Ϣ</a>"
			Case "gtask"	 Rem ��������
				If ChkAble("3,4") Then PageLink=PageLink & "<a href=gtask_assign.asp>��������</a> |"
				If ChkAble(-1) Then PageLink=PageLink & " <a href=mtask_list.asp>�������</a>"
			Case Else			Rem ����
				PageLink=pName & "(��ʱû�пɼ�����)"
		End Select
	End Function
%>