<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
	'Ա����������ļ�
	Call ChkAble(0)
	Dim action, iid, dtKp, strZrr, iGroup, iKpKind, strKpTopic, strKpItem, iKpUPrice, iKpCs, iKpMul, strBz, iZlID, strkpr, strlsh, strTemp,strfz,strGroup
	action="" : iid=0 : dtKp=Now() : strZrr="" : iGroup=0 : iKpKind=0 : strKpItem="" : iKpUPrice=0 : strlsh=""
	iKpCs=1 : iKpMul=-1 : strfz=0 : strBz="" : iZlID=0 : strKpr=Session("username") : strGroup=""
	action=LCase(Request("action"))
	'������⺯�������￪ʼ
	Select Case action
		Case "zrtozykp"		'3���� to ��Ա����
			iKpKind=5		'5Ϊ��Ա����
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
'			If strTemp(2)="" Then
'				iKpUPrice=Request("wtfz")
'			else
'				iKpUPrice=CSng(strTemp(2))
'			End If
			iKpUPrice=Round(Request("kpfz"),1)
			If iKpUPrice="" Then Call JsAlert("�����ֲ���Ϊ��!","")
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")

			strlsh=Trim(Request("kplsh"))
			If strlsh<>"" Then
				strSql="Select * from [mtask] where lsh='"&strlsh&"'"
				Set Rs=xjweb.Exec(strSql,1)
				If Rs.Eof Or Rs.bof Then Call JsAlert("��ˮ���ǲ��������?���ʵһ��!","")
				Dim tempzrr,mtjgshr,dxjgshr,mtsjshr,dxsjshr
				mtjgr=Rs("mtjgr") : dxjgr=Rs("dxjgr") : mtsjr=Rs("mtsjr") : dxsjr=Rs("dxsjr")
				mtjgshr=Rs("mtjgshr") : dxjgshr=Rs("dxjgshr") : mtsjshr=Rs("mtsjshr") : dxsjshr=Rs("dxsjshr")
				tempzrr=Array(mtjgr,dxjgr,mtsjr,dxsjr,mtjgshr,dxjgshr,mtsjshr,dxsjshr)
				For I = Lbound(tempzrr) to Ubound(tempzrr)
					strZrr=tempzrr(i)
					strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
					Set Rs=xjweb.Exec(strSql,1)
					If Not(Rs.Eof Or Rs.Bof) Then
						iGroup=Rs("user_group")
					Rs.Close
					iKpUPrice=CSng(strTemp(2))
					If i>3 Then	iKpUPrice=Round(iKpUPrice/3,2)
					Call kp_add("���� �� ��Ա")
					End If
				Next
			else
				strZrr=Request("kpzrr")
				If strZrr="" Then Call JsAlert("��ѡ������Ա!","")
				strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
				'response.write strsql
				Set Rs=xjweb.Exec(strSql,1)
				If Not(Rs.Eof Or Rs.Bof) Then
					iGroup=Rs("user_group")
				End If
				Rs.Close
				Call kp_add("���� �� ��Ա")
			End If
			If strZrr<>"" Then
				call sendmsg(strZrr, web_info(0), "��������:"&strKpItem&"<br>", "����<b>"&strKpItem&"</b>�����ˣ���ϸ�����뿴�����б�")
			End If
			Call JsAlert("���� �� ��Ա������ӳɹ�!","ygkp_add.asp")

		Case "zrtotsykp"		'4���� to ����Ա����
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("��ѡ������Ա!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=4	'4Ϊ����Ա
			iKpUPrice=Round(Request("kpfz"),1)
			If iKpUPrice="" Then Call JsAlert("�����ֲ���Ϊ��!","")
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")
			Call kp_add("���� �� ����Ա����")
			If strZrr<>"" Then
				call sendmsg(strZrr, web_info(0), "��������:"&strKpItem&"<br>", "����<b>"&strKpItem&"</b>�����ˣ���ϸ�����뿴�����б�")
			End If
			Call JsAlert("���� �� ����Ա����������ӳɹ�!","ygkp_add.asp")

		Case "zrtowgkp"		'4���� to ���ܷ�����Ա����
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("��ѡ������Ա!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=1
			iKpUPrice=Round(Request("kpfz"),1)
			If iKpUPrice="" Then Call JsAlert("�����ֲ���Ϊ��!","")
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")
			Call kp_add("���� �� ������Ա����")
			If strZrr<>"" Then
				call sendmsg(strZrr, web_info(0), "��������:"&strKpItem&"<br>", "����<b>"&strKpItem&"</b>�����ˣ���ϸ�����뿴�����б�")
			End If
			Call JsAlert("���� �� ������Ա������ӳɹ�!","ygkp_add.asp")

		Case "zztotsykp"		'6.�鳤 to ����Ա����
			dim tZrr
			tZrr=Request("kpzrr")
			If tZrr="" Then Call JsAlert("��ѡ������Ա!","")

			iKpKind=4		'4Ϊ����Ա����
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(strTemp(2))
			iKpMul=CInt(strTemp(3))
			strlsh=Request("hglsh")
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")

			tZrr=split(tZrr,"|")
			if Instr(strKpItem,"�ϸ�")>0 Then
				iKpUPrice=iKpUPrice/(ubound(tZrr)+1)  	'����ѡ�����Աƽ��ע���һ
			else
				if ubound(tZrr)>0 Then
					'Call jsalert(ubound(tzrr) & strKpItem,"ygkp_add.asp")
					Call JsAlert("ֻ�кϸ���Ŀ������Ա��ѡ!\n\n������ѡ����Ա!","")
				end if
			end if

			Dim Ttsymjzf, Ttsymttsr, Ttsydxtsr
			Ttsymjzf=""	:	Ttsymttsr="" : Ttsydxtsr=""
			If Instr(strKpItem,"�ϸ�")>0 Then
				strSql="Select * from [mtask] where lsh='"&strlsh&"'"
				Set Rs=xjweb.Exec(strSql,1)
				If Rs.Eof Or Rs.bof Then Call JsAlert("��ˮ���ǲ��������?���ʵһ��!","")
				Set Rs=xjweb.Exec(strSql,1)
				Ttsymjzf=Rs("mjzf")
			End If
			'���γ���
			If Instr(strKpItem,"�ڶ�����Ʒ�ϸ�")>0 Then
			iKpUPrice=iKpUPrice*Ttsymjzf
			End If
			'���γ���
			If Instr(strKpItem,"��������Ʒ�ϸ�")>0 Then
			iKpUPrice=iKpUPrice*Ttsymjzf
			End If

			for i=0 to ubound(tZrr)
				strZrr=tZrr(i)
				strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"

				'response.write strsql
				Set Rs=xjweb.Exec(strSql,1)
				If Not(Rs.Eof Or Rs.Bof) Then
					iGroup=Rs("user_group")
				End If
				Rs.Close
				Call tsykp_Add()
			next
			Call JsAlert("�鳤 �� ����Ա�����ɹ�!","ygkp_add.asp")


		Case "zztozykp"		'7.�鳤 to ��Ա����
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("��ѡ������Ա!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=5		'5Ϊ��Ա����
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(strTemp(2))
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")
			Call kp_add("�鳤 �� ��Ա")
			Call JsAlert("�鳤 �� ��Ա������ӳɹ�!","ygkp_add.asp")


		Case "pgbtotsykp"		'Ʒ�ܲ� to  ���Լ���Ա����
			Dim kpxs, strZrrjs			'����ϵ��,�����˽�ɫ
			dim tsZrr,tsShr, ljjs, ljxs, strbztmp
			tsZrr=Request("kpzrr")
			tsShr=Request("kpsh")
			ljjs=CInt(request("ljjs"))
			ljxs=CSng(request("ljxs"))
			iKpUPrice=CSng(Request("Pgkpfz"))
			If tsZrr="" Then Call JsAlert("��ѡ������Ա!","")
			if ljjs=0 Then Call JsAlert("��ѡ���������!","")
			if iKpUPrice="" Then Call JsAlert("��ֵ����Ϊ��!", "")

			'��ΪƷ�ܲ�һ�ο����漰�ܶ�������ڴ���������������޶�
			Randomize
			iZlID=rnd*99999
			iKpKind=4		'4Ϊ����Ա����
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpMul=CInt(strTemp(3))
			strbztmp=Request("kpbz") & vbcrlf & "�������:" & ljjs & " ϵ��:" & ljxs
			If Request("kpbz")="" Then Call JsAlert("��עΪ��!","")
			'������������
			If tsShr<>"" Then
			strSql="Select [user_group] from [ims_user] where [user_name]='"&tsShr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close
			kpxs=1
			strZrrjs="���"
			strZrr=tsShr
			strBz=strbztmp & vbcrlf & strZrrjs
			Call pgbkp_Add()
			End If

			'�������������
			tsZrr=split(tsZrr,"|")
			if ubound(tsZrr)>0 Then
				iKpUPrice=iKpUPrice/(ubound(tsZrr)+1)  	'����ѡ�����Աƽ��,ע���һ
			end if
			for i=0 to ubound(tsZrr)
				kpxs=1
				strZrrjs="���"
				strZrr=tsZrr(i)
				strBz=strbztmp & vbcrlf & strZrrjs
				strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
				'response.write strsql
				Set Rs=xjweb.Exec(strSql,1)
				If Not(Rs.Eof Or Rs.Bof) Then
					iGroup=Rs("user_group")
				End If
				Rs.Close
				Call pgbkp_Add()
			next
			Call JsAlert("Ʒ�ܲ������Լ���Ա�����ɹ�!","ygkp_add.asp")

		Case "pgbtozykp"		'8.Ʒ�ܲ� to ��Ա����
			strlsh=Trim(Request("kplsh"))
			If strlsh="" Then Call JsAlert("���������ģ�ߵ���ˮ��!","")

			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(Request("Pgkpfz"))
			iKpMul=CInt(strTemp(3))

			'��ΪƷ�ܲ�һ�ο����漰�ܶ�������ڴ���������������޶�
			Randomize
			iZlID=rnd*99999

			iKpKind=5		'5Ϊ��Ա����
			strBz=strlsh&","&Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")

'			If Instr(strKpItem,"��Ʒ�ϸ�")>0 Then
'				strSql="Select * from [mtask] where lsh='"&strlsh&"'"
'				Set Rs=xjweb.Exec(strSql,1)
'				If Rs.Eof Or Rs.bof Then Call JsAlert("��ˮ���ǲ��������?���ʵһ��!","")
'				dim mtjgr, dxjgr, mtsjr, dxsjr, mtshr, dxshr, zmjzf,tempjs
'				mtshr=Rs("mtshr") : dxshr=Rs("dxshr") : zmjzf=Rs("mjzf")
'				mtjgr=Rs("mtjgr") : dxjgr=Rs("dxjgr") :	mtsjr=Rs("mtsjr") : dxsjr=Rs("dxsjr") :
'				mtjgshr=Rs("mtjgshr") : dxjgshr=Rs("dxjgshr") : mtsjshr=Rs("mtsjshr") : dxsjshr=Rs("dxsjshr")
'				'һ�γ���
'				If Instr(strKpItem,"��һ����Ʒ�ϸ�")>0 Then
'				iKpUPrice=iKpUPrice*zmjzf
'				End If
'				'���γ���
'				If Instr(strKpItem,"�ڶ�����Ʒ�ϸ�")>0 Then
'				iKpUPrice=iKpUPrice*zmjzf
'				End If
'				tempzrr=Array(mtjgr,dxjgr,mtsjr,dxsjr,mtjgshr,dxjgshr,mtsjshr,dxsjshr,mtshr,dxshr)
'				tempjs=Array("ģͷ�ṹ","���ͽṹ","ģͷ���","�������","ģͷ�ṹ���","���ͽṹ���","ģͷ������","����������","ģͷ���","�������")
'				For I = Lbound(tempzrr) to Ubound(tempzrr)
'					strZrr=tempzrr(i)
'					If strZrr<>"" Then
'						strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
'						Set Rs=xjweb.Exec(strSql,1)
'						If Not(Rs.Eof Or Rs.Bof) Then
'							iGroup=Rs("user_group")
'						End If
'						Rs.Close
'						strZrrjs=tempjs(i)
'						If InStr(strZrrjs,"�ṹ") Then kpxs=0.3
'						If InStr(strZrrjs,"���") Then kpxs=0.2
'						If InStr(strZrrjs,"���") Then kpxs=0.1
'						Call pgbkp_Add()
'					End If
'				Next
'			Else
				dim sjr, shr
				sjr="" : shr=""
				sjr=Request("kpsj")
				shr=Request("kpsh")
				ljjs=CInt(request("ljjs"))
				ljxs=CSng(request("ljxs"))
				If sjr="" Then Call JsAlert("��ѡ�������!","")
				If shr="" Then Call JsAlert("��ѡ�������!","")
				if ljjs=0 Then Call JsAlert("��ѡ���������!","")
				if ljxs=0.0 Then Call JsAlert("��ѡ�����ϵ��!", "")
				if iKpUPrice="" Then Call JsAlert("��ֵ����Ϊ��!", "")

				sjr=Split(sjr,",")
				shr=Split(shr,",")
				strBz=strBz & vbcrlf & "�������:" & ljjs & " ϵ��:" & ljxs

				'sjr���
				strGroup=""
				For i=0 to ubound(sjr)
					strSql="Select [user_group] from [ims_user] where [user_name]='"&sjr(i)&"'"
					Set Rs=xjweb.Exec(strSql,1)
					If Not(Rs.Eof Or Rs.Bof) Then
						iGroup=Rs("user_group")
					End If
					Rs.Close
					kpxs=1
					If Instr(strGroup,iGroup&",") > 0 Then
						strZrrjs="���2"
					else
						strZrrjs="���"
					End If
					strZrr=sjr(i)
					Call pgbkp_Add()
					strGroup=strGroup&iGroup&","
				next
				'shr���
				strGroup=""
				For i=0 to ubound(shr)
					strSql="Select [user_group] from [ims_user] where [user_name]='"&shr(i)&"'"
					Set Rs=xjweb.Exec(strSql,1)
					If Not(Rs.Eof Or Rs.Bof) Then
						iGroup=Rs("user_group")
					End If
					Rs.Close
					kpxs=1
					If Instr(strGroup,iGroup&",") > 0 Then
						strZrrjs="���2"
					else
						strZrrjs="���"
					End If
					strZrr=shr(i)
					Call pgbkp_Add()
					strGroup=strGroup&iGroup&","
				Next
'			End If
			Call JsAlert("Ʒ�ܲ� �� ����Ա������ӳɹ�!","ygkp_add.asp")

		Case "glbtozrkp"		'9. ���� to ���ο���
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("��ѡ������Ա!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=5		'5Ϊ��Ա����
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("��ѡ��������!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(strTemp(2))
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")
			Call kp_add("�鳤 �� ��Ա")
			Call JsAlert("�鳤 �� ��Ա������ӳɹ�!","ygkp_add.asp")

		Case "ygkpchange"		'���Ŀ�����Ϣ
			iid=Request("id")
			If Not IsNumeric(iid) Then Call JsAlert("�����ȷ��ڽ���!","")
			iid=CLng(iid)
			strfz=Request("kpfz")
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("��עΪ��!","")
			Call kp_Change()
'			Call JsAlert(strfz,"")

		Case Else
			Call JsAlert("action="&action&", ����ϵ����Ա!","")
	End Select

	'����Ա������Ϣ���
	Function tsykp_Add()
		strSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("kp_time")=Now()
			Rs("kp_zrr")=strZrr
			Rs("kp_group")=iGroup
			Rs("kp_kind")=iKpKind
			Rs("kp_topic")=strKpTopic
			Rs("kp_item")=strKpItem
			Rs("kp_uprice")=iKpUPrice
			Rs("kp_cs")=1		'���ǿ�������,ϵͳĬ��Ϊ1
			Rs("kp_mul")=iKpMul
			If strBz<>"" Then Rs("kp_bz")=strBz
			Rs("kp_kpr")=strKpr
		Rs.Update
		Rs.Close
	End Function

	'������Ϣ���
	Function kp_Add(str)
		strSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("kp_time")=Now()
			Rs("kp_zrr")=strZrr
			Rs("kp_group")=iGroup
			Rs("kp_kind")=iKpKind
			Rs("kp_topic")=strKpTopic
			Rs("kp_item")=strKpItem
			Rs("kp_uprice")=iKpUPrice
			Rs("kp_cs")=1		'���ǿ�������,ϵͳĬ��Ϊ1
			Rs("kp_mul")=iKpMul
			If strBz<>"" Then Rs("kp_bz")=strBz
			Rs("kp_kpr")=strKpr
		Rs.Update
		Rs.Close
	End Function

	'Ʒ�ܲ�������Ϣ���
	Function pgbkp_Add()
		Dim strkptime
		strkptime=Request("khsj")
		If strkptime="" Then strkptime=Now()
		strSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("kp_time")=strkptime
			Rs("kp_zrr")=strZrr
			Rs("kp_zrrjs")=strZrrjs
			Rs("kp_group")=iGroup
			Rs("kp_kind")=iKpKind
			Rs("kp_topic")=strKpTopic
			Rs("kp_item")=strKpItem
			Rs("kp_uprice")=iKpUPrice * kpxs
			Rs("kp_cs")=1		'���ǿ�������,ϵͳĬ��Ϊ1
			Rs("kp_mul")=iKpMul
			If strlsh<>"" Then Rs("kp_lsh")=strlsh
			Rs("kp_zlid")=iZlID
			If strBz<>"" Then Rs("kp_bz")=strBz
			Rs("kp_kpr")=strKpr
		Rs.Update
		Rs.Close
	End Function

	'���Ŀ�����Ϣ���
	Function kp_Change()
	Dim strFeedBack, strgzz, strkpjs, strclsh, striPage, strkptime, strkpitem
	strkptime=Request("kpsj")
	strZrr=Trim(Request("zrr"))
	strkpitem = trim(request("kpitem"))
	strkpjs = trim(request("kpjs"))
	strclsh = trim(request("kplsh"))
	strgzz =request("kpgzz")
	striPage =request("ipage")
	strFeedBack=""
	If strZrr<>"" Then strFeedBack="zrr="&strZrr
	If strkpitem<>"" Then strFeedBack="kpitem="&strkpitem&"&"&strFeedBack
	If strgzz<>"" Then strFeedBack="gzz="&strgzz&"&"&strFeedBack
	If strkpjs<>"" Then strFeedBack="kpjs="&strkpjs&"&"&strFeedBack
	If strclsh<>"" Then strFeedBack="clsh="&strclsh&"&"&strFeedBack
	If striPage<>"0" Then strFeedBack="iPage="&striPage&"&"&strFeedBack
	If strFeedBack<>"" Then strFeedBack="?"&strFeedBack

		'���ID���Ƿ����
		Set Rs=xjweb.Exec("select * from [kp_jsb] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("�˼���������Ϣ�����Ѿ�ɾ��!","ygkp_list.asp"&strFeedBack)
			Rs.Close
			Exit Function
		End If
		Rs.Close
			strSql="select * from [kp_jsb] where id=" & iid
			Call xjweb.Exec("",-1)
			Rs.open strSql,conn,1,3
				If strBz<>"" Then Rs("kp_bz")=strBz
				If strfz<>"" Then Rs("kp_uprice")=strfz
				If strkptime<>"" Then Rs("kp_time")=strkptime
			Rs.update
			Rs.close

		Call JsAlert("Ա���������ĳɹ�","ygkp_list.asp"&strFeedBack)
	End Function
%>
