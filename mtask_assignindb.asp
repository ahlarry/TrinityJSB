<!--#include file="include/conn.asp"-->
<%
'11:22 2007-5-30-������
	'���ļ�ֻ�����������������
	Call ChkPageAble("3,4")
	dim strfplr, strzrr, strlsh, strSql2, Rs2, strhth
	strSql2=""
	strfplr=request("fplr")
	strzrr=request("zrr")
	strlsh=request("lsh")
	strhth=request("hth")
	if strfplr="" or (strzrr="" and instr(strfplr,"��ʼ")>0) or (strlsh="" and strhth="") then
		Call JsAlert("������������Ϣ����! \n��������Ϊ�ջ�û��������!","")
	end if
	dim bok
	bok=true
	if strlsh<>"" Then
		strSql2="select * from [mtask] where lsh='"&strlsh&"'"
		Set Rs2=xjweb.Exec(strSql2,1)
	else
		strSql2="select * from [jsdb] where hth='"&strhth&"'"
		Set Rs2=xjweb.Exec(strSql2,1)
	End if
	strSql=""
	select case strfplr
		case "��ʼģͷ�ṹ"
			strSql="update [mtask] set mtjgr='"&strzrr&"', mtjgks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷ�ṹ", strzrr, now())

		case "��ʼ���ͽṹ"
			strSql="update [mtask] set dxjgr='"&strzrr&"', dxjgks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "���ͽṹ", strzrr, now())

		case "��ʼ�󹲼��ṹ"
			strSql="update [mtask] set gjjgr='"&strzrr&"', gjjgks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�󹲼��ṹ", strzrr, now())

		case "��ʼȫ�׽ṹ"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtjgr='"&strzrr&"', mtjgks='"&now()&"', dxjgr='"&strzrr&"', gjjgr='"&strzrr&"', gjjgks='"&now()&"', dxjgks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ�ṹ", strzrr, now())
				call taskstart(strlsh, "���ͽṹ", strzrr, now())
				call taskstart(strlsh, "�󹲼��ṹ", strzrr, now())
			else
				strSql="update [mtask] set mtjgr='"&strzrr&"', mtjgks='"&now()&"', dxjgr='"&strzrr&"', dxjgks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ�ṹ", strzrr, now())
				call taskstart(strlsh, "���ͽṹ", strzrr, now())
			end if

		case "��ʼģͷ�ṹȷ��"
			strSql="update [mtask] set mtjgshr='"&strzrr&"', mtjgshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷ�ṹȷ��", strzrr, now())

		case "��ʼ���ͽṹȷ��"
			strSql="update [mtask] set dxjgshr='"&strzrr&"', dxjgshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "���ͽṹȷ��", strzrr, now())

		case "��ʼ�󹲼��ṹȷ��"
			strSql="update [mtask] set gjjgshr='"&strzrr&"', gjjgshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�󹲼��ṹȷ��", strzrr, now())

		case "��ʼȫ�׽ṹȷ��"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtjgshr='"&strzrr&"', mtjgshks='"&now()&"', dxjgshr='"&strzrr&"', gjjgshr='"&strzrr&"', dxjgshks='"&now()&"', gjjgshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ�ṹȷ��", strzrr, now())
				call taskstart(strlsh, "���ͽṹȷ��", strzrr, now())
				call taskstart(strlsh, "�󹲼��ṹȷ��", strzrr, now())
			else
				strSql="update [mtask] set mtjgshr='"&strzrr&"', mtjgshks='"&now()&"', dxjgshr='"&strzrr&"', dxjgshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ�ṹȷ��", strzrr, now())
				call taskstart(strlsh, "���ͽṹȷ��", strzrr, now())
			end if

		case "��ʼģͷ���"
			strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷ���", strzrr, now())

		case "��ʼ�������"
			strSql="update [mtask] set dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�������", strzrr, now())

		case "��ʼ�󹲼����"
			strSql="update [mtask] set gjsjr='"&strzrr&"', gjsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�󹲼����", strzrr, now())

		case "��ʼȫ�����"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', gjsjr='"&strzrr&"', gjsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ���", strzrr, now())
				call taskstart(strlsh, "�������", strzrr, now())
				call taskstart(strlsh, "�󹲼����", strzrr, now())
			else
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ���", strzrr, now())
				call taskstart(strlsh, "�������", strzrr, now())
			end if

		case "��ʼģͷ���ȷ��"
			strSql="update [mtask] set mtsjshr='"&strzrr&"', mtsjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷ���ȷ��", strzrr, now())

		case "��ʼ�������ȷ��"
			strSql="update [mtask] set dxsjshr='"&strzrr&"', dxsjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�������ȷ��", strzrr, now())

		case "��ʼ�󹲼����ȷ��"
			strSql="update [mtask] set gjsjshr='"&strzrr&"', gjsjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�󹲼����ȷ��", strzrr, now())

		case "��ʼȫ�����ȷ��"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtsjshr='"&strzrr&"', mtsjshks='"&now()&"', dxsjshr='"&strzrr&"', gjsjshr='"&strzrr&"', dxsjshks='"&now()&"', gjsjshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ���ȷ��", strzrr, now())
				call taskstart(strlsh, "�������ȷ��", strzrr, now())
				call taskstart(strlsh, "�󹲼����ȷ��", strzrr, now())
			else
				strSql="update [mtask] set mtsjshr='"&strzrr&"', mtsjshks='"&now()&"', dxsjshr='"&strzrr&"', dxsjshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ���ȷ��", strzrr, now())
				call taskstart(strlsh, "�������ȷ��", strzrr, now())
			end if

		case "��ʼģͷ���"
			strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷ���", strzrr, now())

		case "��ʼ�������"
			strSql="update [mtask] set dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�������", strzrr, now())

		case "��ʼ�󹲼����"
			strSql="update [mtask] set gjshr='"&strzrr&"', gjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�󹲼����", strzrr, now())

		case "��ʼȫ�����"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', gjshr='"&strzrr&"', gjshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ���", strzrr, now())
				call taskstart(strlsh, "�������", strzrr, now())
				call taskstart(strlsh, "�󹲼����", strzrr, now())
			else
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ���", strzrr, now())
				call taskstart(strlsh, "�������", strzrr, now())
			end if

		case "ȫ�׹������"
			strSql="update [mtask] set mtgysjr='"&strzrr&"', mtgysjks='"&now()&"', mtgysjjs='"&now()&"', dxgysjr='"&strzrr&"', dxgysjks='"&now()&"', dxgysjjs='"&now()&"', gjgysjr='"&strzrr&"', gjgysjks='"&now()&"', gjgysjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
		case "ģͷ�������"
			strSql="update [mtask] set  mtgysjr='"&strzrr&"', mtgysjks='"&now()&"', mtgysjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
		case "���͹������"
			strSql="update [mtask] set dxgysjr='"&strzrr&"', dxgysjks='"&now()&"', dxgysjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
		case "�����������"
			strSql="update [mtask] set gjgysjr='"&strzrr&"', gjgysjks='"&now()&"', gjgysjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
		case "ȫ�׹������"
			strSql="update [mtask] set mtgyshr='"&strzrr&"', mtgyshks='"&now()&"', mtgyshjs='"&now()&"', dxgyshr='"&strzrr&"', dxgyshks='"&now()&"', dxgyshjs='"&now()&"', gjgyshr='"&strzrr&"', gjgyshks='"&now()&"', gjgyshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
		case "ģͷ�������"
			strSql="update [mtask] set mtgyshr='"&strzrr&"', mtgyshks='"&now()&"', mtgyshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
		case "���͹������"
			strSql="update [mtask] set dxgyshr='"&strzrr&"', dxgyshks='"&now()&"', dxgyshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
		case "�����������"
			strSql="update [mtask] set gjgyshr='"&strzrr&"', gjgyshks='"&now()&"', gjgyshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)

		case "��ʼģͷBOM"
			strSql="update [mtask] set mtbomr='"&strzrr&"', mtbomks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷBOM", strzrr, now())

		case "��ʼ����BOM"
			strSql="update [mtask] set dxbomr='"&strzrr&"', dxbomks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "����BOM", strzrr, now())

		case "��ʼȫ��BOM"
			strSql="update [mtask] set mtbomr='"&strzrr&"', mtbomks='"&now()&"', dxbomr='"&strzrr&"', dxbomks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷBOM", strzrr, now())
			call taskstart(strlsh, "����BOM", strzrr, now())

		case "��ʼģͷ����"
			strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷ����", strzrr, now())

		case "��ʼ���͸���"
			strSql="update [mtask] set dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "���͸���", strzrr, now())

		case "��ʼ�󹲼�����"
			strSql="update [mtask] set gjsjr='"&strzrr&"', gjsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�󹲼�����", strzrr, now())

		case "��ʼȫ�׸���"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', gjsjr='"&strzrr&"', gjsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ����", strzrr, now())
				call taskstart(strlsh, "���͸���", strzrr, now())
				call taskstart(strlsh, "�󹲼�����", strzrr, now())
			else
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ����", strzrr, now())
				call taskstart(strlsh, "���͸���", strzrr, now())
			end if

		case "��ʼģͷ����"
			strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "ģͷ����", strzrr, now())

		case "��ʼ���͸���"
			strSql="update [mtask] set dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "���͸���", strzrr, now())

		case "��ʼ�󹲼�����"
			strSql="update [mtask] set gjshr='"&strzrr&"', gjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "�󹲼�����", strzrr, now())

		case "��ʼȫ�׸���"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', gjshr='"&strzrr&"', gjshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ����", strzrr, now())
				call taskstart(strlsh, "���͸���", strzrr, now())
				call taskstart(strlsh, "�󹲼�����", strzrr, now())
			else
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "ģͷ����", strzrr, now())
				call taskstart(strlsh, "���͸���", strzrr, now())
			end if

		case "����ģͷ�ṹ"
			strSql="update [mtask] set mtjgjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷ�ṹ")

		case "�������ͽṹ"
			strSql="update [mtask] set dxjgjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "���ͽṹ")

		case "�����󹲼��ṹ"
			strSql="update [mtask] set gjjgjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�󹲼��ṹ")

		case "����ȫ�׽ṹ"
			if not(isnull(Rs2("gjjgks"))) Then
				strSql="update [mtask] set mtjgjs='"&now()&"', dxjgjs='"&now()&"', gjjgjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ�ṹ")
				call taskend(strlsh, "���ͽṹ")
				call taskend(strlsh, "�󹲼��ṹ")
			else
				strSql="update [mtask] set mtjgjs='"&now()&"', dxjgjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ�ṹ")
				call taskend(strlsh, "���ͽṹ")
			End if

		case "����ģͷ�ṹȷ��"
			strSql="update [mtask] set mtjgshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷ�ṹȷ��")

		case "�������ͽṹȷ��"
			strSql="update [mtask] set dxjgshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "���ͽṹȷ��")

		case "�����󹲼��ṹȷ��"
			strSql="update [mtask] set gjjgshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�󹲼��ṹȷ��")

		case "����ȫ�׽ṹȷ��"
			if not(isnull(Rs2("gjjgshks"))) Then
				strSql="update [mtask] set mtjgshjs='"&now()&"', dxjgshjs='"&now()&"', gjjgshjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ�ṹȷ��")
				call taskend(strlsh, "���ͽṹȷ��")
				call taskend(strlsh, "�󹲼��ṹȷ��")
			else
				strSql="update [mtask] set mtjgshjs='"&now()&"', dxjgshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ�ṹȷ��")
				call taskend(strlsh, "���ͽṹȷ��")
			End if

		case "����ģͷ���"
			strSql="update [mtask] set mtsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷ���")

		case "�����������"
			strSql="update [mtask] set dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�������")

		case "�����󹲼����"
			strSql="update [mtask] set gjsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�󹲼����")

		case "����ȫ�����"
			if not(isnull(Rs2("gjsjks"))) Then
				strSql="update [mtask] set mtsjjs='"&now()&"', gjsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ���")
				call taskend(strlsh, "�������")
				call taskend(strlsh, "�󹲼����")
			else
				strSql="update [mtask] set mtsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ���")
				call taskend(strlsh, "�������")
			End if

		case "����ģͷ���ȷ��"
			strSql="update [mtask] set mtsjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷ���ȷ��")

		case "�����������ȷ��"
			strSql="update [mtask] set dxsjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�������ȷ��")

		case "�����󹲼����ȷ��"
			strSql="update [mtask] set gjsjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�󹲼����ȷ��")

		case "����ȫ�����ȷ��"
			if not(isnull(Rs2("gjsjshks"))) Then
				strSql="update [mtask] set mtsjshjs='"&now()&"', dxsjshjs='"&now()&"', gjsjshjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ���ȷ��")
				call taskend(strlsh, "�������ȷ��")
				call taskend(strlsh, "�󹲼����ȷ��")
			else
				strSql="update [mtask] set mtsjshjs='"&now()&"', dxsjshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ���ȷ��")
				call taskend(strlsh, "�������ȷ��")
			End if

		case "����ģͷ���"
			strSql="update [mtask] set mtshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷ���")

		case "�����������"
			strSql="update [mtask] set dxshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�������")

		case "�����󹲼����"
			strSql="update [mtask] set gjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�󹲼����")

		case "����ȫ�����"
			if not(isnull(Rs2("gjshks"))) Then
				strSql="update [mtask] set mtshjs='"&now()&"', gjshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ���")
				call taskend(strlsh, "�������")
				call taskend(strlsh, "�󹲼����")
			else
				strSql="update [mtask] set mtshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ���")
				call taskend(strlsh, "�������")
			End if

		case "����ģͷBOM"
			strSql="update [mtask] set mtbomjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷBOM")

		case "��������BOM"
			strSql="update [mtask] set dxbomjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "����BOM")

		case "����ȫ��BOM"
			strSql="update [mtask] set mtbomjs='"&now()&"', dxbomjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷBOM")
			call taskend(strlsh, "����BOM")

		case "����ģͷ����"
			strSql="update [mtask] set mtsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷ����")

		case "�������͸���"
			strSql="update [mtask] set dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "���͸���")

		case "�����󹲼�����"
			strSql="update [mtask] set gjsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�󹲼�����")

		case "����ȫ�׸���"
			if not(isnull(Rs2("gjsjks"))) Then
				strSql="update [mtask] set mtsjjs='"&now()&"', gjsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ����")
				call taskend(strlsh, "���͸���")
				call taskend(strlsh, "�󹲼�����")
			else
				strSql="update [mtask] set mtsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ����")
				call taskend(strlsh, "���͸���")
			End if

		case "����ģͷ����"
			strSql="update [mtask] set mtshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "ģͷ����")

		case "�������͸���"
			strSql="update [mtask] set dxshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "���͸���")

		case "�����󹲼�����"
			strSql="update [mtask] set gjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "�󹲼�����")

		case "����ȫ�׸���"
			if not(isnull(Rs2("gjsjks"))) Then
				strSql="update [mtask] set mtshjs='"&now()&"', gjshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ����")
				call taskend(strlsh, "���͸���")
				call taskend(strlsh, "�󹲼�����")
			else
				strSql="update [mtask] set mtshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "ģͷ����")
				call taskend(strlsh, "���͸���")
			End if

		case "��������"
			strSql="update [mtask] set psjl='"&request("psjl")&"', fsr='"&strzrr&"', fsjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call sendmsg("����", web_info(0), "ģ��ȫ�׽���", "<a href=mtask_display.asp?s_lsh="&strlsh&" target=""_blank"">��ˮ�� <b>"&strlsh&"</b> �ѽ�����������׼�豾ģ��ȫ�׽�����</a>")
			Call JsAlert("��ˮ�� ��" & strlsh & "�� �������鳤���䲿�����!","mtask_assign.asp")

		case "ȫ�׽���"
			if not bmtaskend("sjjssj", strlsh) then
				strSql="update [mtask] set sjjssj='"&request("psd")&" "& request("pst") &"', psjl='"&request("zrpsjl")&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call fentodb(strlsh)
			end if
			Call JsAlert("��ˮ�� ��" & strlsh & "�� ������ȫ�����!","mtask_assign.asp")

		case "��ʼ�����������"
			strSql="update [jsdb] set sjr='"&strzrr&"', sjkssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strhth, "�����������", strzrr, now())
		case "���������������"
			strSql="update [jsdb] set sjjssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strhth, "�����������")

		case "��ʼ�����������"
			strSql="update [jsdb] set shr='"&strzrr&"', shkssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strhth, "�����������", strzrr, now())
		case "���������������"
			strSql="update [jsdb] set shjssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strhth, "�����������")
			call jsdbfz(strhth,Rs2("khmc"),Rs2("rwnr"),Rs2("jcf"),Rs2("jhjssj"),Rs2("sjr"),Rs2("sjjssj"),Rs2("shr"),Rs2("shjssj"))
		case else
			bok=false
			Call JsAlert("������������ݲ���ϵ����Ա:\n\n"& strfplr & "\n\n (ϵͳ�쳣)����ϵ����Ա","mtask_assign.asp")
	end select
	if bok and strfplr<>"��������" and strfplr<>"ȫ�׽���" then
		Call JsAlert("���������ɹ�!","mtask_assign.asp?s_lsh="&strlsh&"&s_hth="&strhth&"")
	end if

function taskstart(lsh, rwlr, zrr, kssj)
	strSql="insert into [mtask_cur] (lsh, rwlr, zrr, kssj) values ('"&lsh&"', '"&rwlr&"', '"&zrr&"','"&kssj&"')"
	call xjweb.Exec(strSql, 0)
end function

function taskend(lsh, rwlr)
	strSql="delete from [mtask_cur] where lsh='"&lsh&"' and rwlr='"&rwlr&"'"
	call xjweb.Exec(strSql, 0)
end function

function bmtaskend(trs, lsh)		'��ֹ��ֵ�ٴ����
	bmtaskend=true
	strSql="select "&trs&" from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	if isnull(rs(trs)) then bmtaskend=false
	rs.close
end function

Function FenToDB(lsh)
	'����ֵд���ֵ��
	Dim mjfz, mtfz, dxfz, gjfz, bomfz, ijgbl, isjbl, ishbl, ifgbl, ifgshbl, ifcbl, ijc, ijc2, imtjgbl, idxjgbl, ijgshbl, iljshbl, iwcsj, mtgjf, dxgjf, ssgjf, qbfgjf, qgjf, hgjf
	Dim igysjxs, igysjsh, igyfcxs, igyfcsh, igyfgxs, igyfgsh, iGroup, tmpSql, tmpRs
	mtgjf=0 : dxgjf=0
	'ijc===���ͷ�ֵ
	strSql="select * from [c_fzbl]"
	set rs=xjweb.Exec(strSql, 1)
	imtjgbl=CSng(rs("mtjgbl"))
	idxjgbl=CSng(rs("dxjgbl"))
	ijgshbl=CSng(rs("jgshbl"))
	iljshbl=CSng(rs("ljshbl"))
	ishbl=CSng(rs("shbl"))
	ifgbl=CSng(rs("fgbl"))
	ifgshbl=CSng(rs("fgshbl"))
	ifcbl=CSng(rs("fcbl"))
	igysjxs=CSng(rs("gysjxs"))
	igysjsh=CSng(rs("gysjsh"))
	igyfcxs=CSng(rs("gyfcxs"))
	igyfcsh=CSng(rs("gyfcsh"))
	igyfgxs=CSng(rs("gyfgxs"))
	igyfgsh=CSng(rs("gyfgsh"))
	rs.close

	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	mjfz=rs("mjzf")
	gjfz=Rs("gjzf")
	ssgjf=NullToNum(Rs("ssgj"))
	qbfgjf=NullToNum(Rs("qbfgj"))
	qgjf=NullToNum(Rs("qgj"))
	hgjf=NullToNum(Rs("hgj"))
	if NullToNum(Rs("mtjgbl"))<>0 Then imtjgbl=Rs("mtjgbl")/100
	if NullToNum(Rs("dxjgbl"))<>0 Then idxjgbl=Rs("dxjgbl")/100
	select case ssgjf&qbfgjf&qgjf&hgjf
		Case "0000"			'����08�湲���Ʒ�ģʽ
			'ֻ����Ӳǰ�����ķ�ֵ�Ų��ּӵ�ģͷ���ּӵ�������
			if Rs("gjfs")="3" and Rs("qhgj")="1" Then
				mtfz=Rs("mjzf")*Rs("mtbl")/100
				dxfz=Rs("mjzf")*(100-Rs("mtbl"))/100
			End if
			'��Ӳ�󹲼��ķ�ֵ�����ӵ��󹲼�����
			If Rs("gjfs")="3" and Rs("qhgj")="2" Then
				mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100
				dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
			'�������������й������ȫ�ӵ�ģͷ
			If (not (Rs("gjfs")="3")) Then
				mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + gjfz
				dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
		Case Else		'09�湲���Ʒ�ģʽ
			If qgjf<>0 Then
				mtgjf=qgjf*Rs("mtbl")/100
				dxgjf=qgjf-mtgjf
			End If
			mtgjf=mtgjf+ssgjf+qbfgjf
			mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + mtgjf
			dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100 + dxgjf
	end select

	bomfz=rs("bomzf")
	ijc2=datediff("d",rs("sjjssj"),rs("jhjssj"))
	ijc=0

	Dim ijgsj, isj
	If IsNull(rs("jhjgsj")) Then
		isj=INT(datediff("d", rs("jhkssj"), rs("jhjssj"))/2)
		ijgsj=dateadd("d",isj,rs("jhkssj"))
	else
		ijgsj=rs("jhjgsj")
	End if
'������ģͷ�Ͷ��ͽṹ����ʱ��ȡ����������һ��
	select case rs("rwlr")
		case "���"
			'�ṹ
			if not(isnull(rs("mtjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ�ṹ',"&Round(mtfz*imtjgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtjgr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If datediff("n", rs("mtjgjs"), rs("dxjgjs"))<0 Then
					iwcsj=rs("mtjgjs")
				else
					iwcsj=rs("dxjgjs")
				End If
				Call Ddkp(rs("mtjgr"), iwcsj, ijgsj, rs("lsh"), "ģͷ�ṹ")
			end if
			if not(isnull(rs("dxjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���ͽṹ',"&Round(dxfz*idxjgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxjgr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If datediff("n", rs("mtjgjs"), rs("dxjgjs"))<0 Then
					iwcsj=rs("mtjgjs")
				else
					iwcsj=rs("dxjgjs")
				End If
				Call Ddkp(rs("dxjgr"), iwcsj, ijgsj, rs("lsh"), "���ͽṹ")
			end if
			if  not(isnull(rs("gjjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�󹲼��ṹ',"&Round(hgjf*idxjgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjjgr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("gjjgr"), rs("gjjgjs"), ijgsj, rs("lsh"), "�󹲼��ṹ")
			end if

			'���
			if not(isnull(rs("mtsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���',"&Round(mtfz*(1-imtjgbl),1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("mtsjr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "ģͷ���")
			end if
			if not(isnull(rs("dxsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������',"&Round(dxfz*(1-idxjgbl),1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("dxsjr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "�������")
			end if
			if not(isnull(rs("gjsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�󹲼����',"&Round(hgjf*(1-idxjgbl),1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("gjsjr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "�󹲼����")
			end if

			'��ģͷ�Ͷ����зֱ������ṹ��������
			if not(isnull(rs("mtjgshr"))) then 'ģͷ�ṹ���
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ�ṹȷ��',"&Round(mtfz*imtjgbl*ijgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtjgshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If datediff("n", rs("mtjgshjs"), rs("dxjgshjs"))<0 Then
					iwcsj=rs("mtjgshjs")
				else
					iwcsj=rs("dxjgshjs")
				End If
				Call Ddkp(rs("mtjgshr"), iwcsj, ijgsj, rs("lsh"), "ģͷ�ṹȷ��")
			end if

			if not(isnull(rs("mtsjshr"))) then 'ģͷ������
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ȷ��',"&Round(mtfz*(1-imtjgbl)*iljshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtsjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("mtsjshr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "ģͷ���ȷ��")
			end if

			if not(isnull(rs("dxjgshr"))) then '���ͽṹ���
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���ͽṹȷ��',"&Round(dxfz*idxjgbl*ijgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxjgshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If datediff("n", rs("mtjgshjs"), rs("dxjgshjs"))<0 Then
					iwcsj=rs("mtjgshjs")
				else
					iwcsj=rs("dxjgshjs")
				End If
				Call Ddkp(rs("dxjgshr"), iwcsj, ijgsj, rs("lsh"), "���ͽṹȷ��")
			end if

			if not(isnull(rs("dxsjshr"))) then '����������
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ȷ��',"&Round(dxfz*(1-idxjgbl)*iljshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxsjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("dxsjshr"), rs("sjjssj"),  rs("jhjssj"), rs("lsh"), "�������ȷ��")
			end if

			if not(isnull(rs("gjjgshr"))) then '�󹲼��ṹ���
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�󹲼��ṹȷ��',"&Round(hgjf*idxjgbl*ijgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjjgshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("gjjgshr"), rs("gjjgshjs"), ijgsj, rs("lsh"), "�󹲼��ṹȷ��")
			end if

			if not(isnull(rs("gjsjshr"))) then '�󹲼�������
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�󹲼����ȷ��',"&Round(hgjf*(1-idxjgbl)*iljshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjsjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("gjsjshr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "�󹲼����ȷ��")
			end if

			'���
			if not(isnull(rs("mtshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���',"&Round(mtfz*ishbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("mtshr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "ģͷ���")
			end if
			if not(isnull(rs("dxshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������',"&Round(dxfz*ishbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("dxshr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "�������")
			end if
			if not(isnull(rs("gjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�󹲼����',"&Round(hgjf*ishbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				Call Ddkp(rs("gjshr"), rs("sjjssj"), rs("jhjssj"), rs("lsh"), "�󹲼����")
			end if

			'Bom
			if not(isnull(rs("mtbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷBOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','����BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			'����
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ�������',"&Round(mtfz*igysjxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹������',"&Round(dxfz*igysjxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�����������',"&Round(hgjf*igysjxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("mtgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ�������',"&Round(mtfz*igysjsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹������',"&Round(dxfz*igysjsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�����������',"&Round(hgjf*igysjsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
		case "����"
			'����
			if not(isnull(rs("mtsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ����',"&Round(mtfz*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("mtsjr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "ģͷ����")
			end if
			if not(isnull(rs("dxsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͸���',"&Round(dxfz*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("dxsjr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "���͸���")
			end if
			if not(isnull(rs("gjsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','��������',"&Round(gjfz*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("gjsjr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "��������")
			end if
			'���
			if not(isnull(rs("mtshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���',"&Round(mtfz*ifgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("mtshr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "ģͷ�������")
			end if
			if not(isnull(rs("dxshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������',"&Round(dxfz*ifgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("dxshr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "���͸������")
			end if
			if not(isnull(rs("gjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������',"&Round(gjfz*ifgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("dxshr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "�����������")
			end if
			'BOM
			if not(isnull(rs("mtbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷBOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','����BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			'����
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո���',"&Round(mtfz*igyfgxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո���',"&Round(dxfz*igyfgxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո���',"&Round(hgjf*igyfgxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("mtgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո������',"&Round(mtfz*igyfgsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո������',"&Round(dxfz*igyfgsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո������',"&Round(hgjf*igyfgsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
		case "����"
			'����
			if not(isnull(rs("mtshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ����',"&Round(mtfz*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("mtshr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "ģͷ����")
			end if
			if not(isnull(rs("dxshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͸���',"&Round(dxfz*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("dxshr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "���͸���")
			end if
			if not(isnull(rs("gjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','��������',"&Round(gjfz*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
				If ijc2<>0 Then Call ygkptodb(rs("gjshr"), ijc2, 1.5, rs("lsh"), rs("lsh"), "��������")
			end if
			'BOM
			if not(isnull(rs("mtbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷBOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','����BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			'����
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո���',"&Round(mtfz*igyfcxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո���',"&Round(dxfz*igyfcxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո���',"&Round(hgjf*igyfcxs,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("mtgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո������',"&Round(mtfz*igyfcsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո������',"&Round(dxfz*igyfcsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո������',"&Round(hgjf*igyfcsh,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
	end select
	rs.close
end function

Function ygkptodb(zrr, tt, iprice, strlsh, lsh, zrrjs)
	dim tmpSql, tmpRs, iGroup, iKPKind, strKpTopic, strKpItem, ikpmul, strbz,isjjssj
	strZrr=zrr
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

	tmpSql="select * from [mtask] where lsh='"&strlsh&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		isjjssj=tmpRs("sjjssj")
	End If
	tmpRs.Close

	iKpKind=5		'5Ϊ��Ա����
	strKpTopic="�������"
	If tt>0 Then
		strKpItem="��ǰ"
	Else
		strKpItem="�ӳ�"
	End If

	If tt>0 Then
		iKpMul=1
	Else
		iKpMul=-1
	End If
	strBz=lsh

	tmpSql="select * from [kp_jsb]"
	Call xjweb.Exec("",-1)
	set tmpRs=Server.CreateObject("adodb.recordset")
	tmpRs.open tmpSql,conn,1,3
	tmpRs.AddNew
		tmpRs("kp_time")=isjjssj
		tmpRs("kp_zrr")=strZrr
		tmpRs("kp_zrrjs")=zrrjs
		tmpRs("kp_group")=iGroup
		tmpRs("kp_kind")=iKpKind
		tmpRs("kp_topic")=strKpTopic
		tmpRs("kp_item")=strKpItem
		tmpRs("kp_uprice")=iPrice
		tmpRs("kp_cs")=1		'���ǿ�������,ϵͳĬ��Ϊ1
		tmpRs("kp_mul")=iKpMul
		tmpRs("kp_bz")=strBz
		tmpRs("kp_lsh")=strlsh
		tmpRs("kp_kpr")=session("userName")
	tmpRs.Update
	tmpRs.Close
End Function

Function Ddkp(zrr,wcsj,jgsj,lsh,zrrjs)
	'�ṹ,��Ƶ���������ǰ������
	'Ddkp(������,���ʱ��,�������ʱ��,��ˮ��,�����ɫ)
	'��ǰ��1.5�֣����ڿ�2.5��
	dim tmpSql, tmpRs, ikp, iGroup, iPrice, iKPKind, strKpTopic, strKpItem, ikpmul, strbz, izz
	If datediff("d",wcsj,jgsj)>0 Then
		ikp=1
		strKpItem="��ǰ"
		iKpMul=1
		iPrice=1.5*datediff("d",wcsj,jgsj)
	else
		If datediff("d",wcsj,jgsj)<0 Then
			ikp=1
			strKpItem="�ӳ�"
			iKpMul=-1
			iPrice=2.5*datediff("d",jgsj,wcsj)
		else
			ikp=0
		End If
	End If
	If (InStr(zrrjs, "���") > 0) or (InStr(zrrjs, "ȷ��") > 0) Then iPrice = iPrice * 0.5

If ikp=1 Then
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&zrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

	iKpKind=5		'5Ϊ��Ա����
	strKpTopic="�������"
	strBz=lsh
	tmpSql="select * from [kp_jsb]"
	Call xjweb.Exec("",-1)
	set tmpRs=Server.CreateObject("adodb.recordset")
	tmpRs.open tmpSql,conn,1,3
	tmpRs.AddNew
		tmpRs("kp_time")=wcsj
		tmpRs("kp_zrr")=zrr
		tmpRs("kp_zrrjs")=zrrjs
		tmpRs("kp_group")=iGroup
		tmpRs("kp_kind")=iKpKind
		tmpRs("kp_topic")=strKpTopic
		tmpRs("kp_item")=strKpItem
		tmpRs("kp_uprice")=iPrice
		tmpRs("kp_cs")=1		'���ǿ�������,ϵͳĬ��Ϊ1
		tmpRs("kp_mul")=iKpMul
		tmpRs("kp_bz")=strBz
		tmpRs("kp_lsh")=strlsh
		tmpRs("kp_kpr")=session("userName")
	tmpRs.Update
	tmpRs.Close

	'������Ӧ�鳤
	izz=""
	If datediff("d",wcsj,jgsj)>0 Then
		iPrice=1*datediff("d",wcsj,jgsj)
	else
		iPrice=2*datediff("d",jgsj,wcsj)
	End If
	tmpSql="select * from [ims_user] where mid(user_able,4,1)>0 and user_group="&iGroup
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		izz=tmpRs("user_name")
	End If
	tmpRs.Close
	if izz<>"" Then
		tmpSql="select * from [kp_jsb] where kp_zrrjs like '%�鳤%' and Instr('��ǰ�ӳ�',kp_item)>0 and kp_lsh='"&strlsh&"' and kp_zrr='"&izz&"'"
		Call xjweb.Exec("",-1)
		tmpRs.open tmpSql,conn,1,3
		If tmpRs.Eof Or tmpRs.Bof Then
			tmpRs.addnew
			tmpRs("kp_time")=wcsj
			tmpRs("kp_zrr")=izz
			tmpRs("kp_zrrjs")="�鳤"
			tmpRs("kp_group")=iGroup
			tmpRs("kp_kind")=3
			tmpRs("kp_topic")=strKpTopic
			tmpRs("kp_item")=strKpItem
			tmpRs("kp_uprice")=iPrice
			tmpRs("kp_cs")=1		'���ǿ�������,ϵͳĬ��Ϊ1
			tmpRs("kp_mul")=iKpMul
			tmpRs("kp_bz")=strBz
			tmpRs("kp_lsh")=strlsh
			tmpRs("kp_kpr")=session("userName")
			tmpRs.update
		End If
		tmpRs.close
	End If
End If
End Function

function jsdbfz(hth,khmc,rwnr,fz,jhsj,sjr,sjjs,shr,shjs)
	Dim tmpSql, tmpRs, iGroup
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&sjr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

		strSql="select * from [ftask]"
		call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		rs.addnew
			rs("rwlx")="�����������"
			rs("rwlr")=hth&khmc&":"&rwnr
			rs("zrr")=sjr
			rs("xz")=iGroup
			rs("zf")=fz
			rs("jssj")=sjjs
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update

	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&shr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

		rs.addnew
			rs("rwlx")="�����������"
			rs("rwlr")=hth&khmc&":"&rwnr
			rs("zrr")=shr
			rs("xz")=iGroup
			rs("zf")=Round(fz/3,1)
			rs("jssj")=shjs
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update
		rs.close
		If datediff("d",sjjs,jhsj)<>0 Then Call ygkptodb(sjr, datediff("d",sjjs,jhsj), 1.5, hth, hth, "�����������")
		If datediff("d",shjs,jhsj)<>0 Then Call ygkptodb(shr, datediff("d",shjs,jhsj), 0.8, hth, hth, "�����������")
End Function
%>