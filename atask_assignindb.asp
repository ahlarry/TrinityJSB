<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble("4,6")
	'���ļ�ֻ������丨����������
'	Call JsAlert("�û��������벻��ȷ,���ʵ������!","")
	Dim strfplr, strzrr, strlsh
	strfplr=Request("fplr")
	strzrr=Request("zrr")
	strlsh=Request("lsh")
	If strfplr="" Or (strzrr="" And InStr(strfplr,"��ʼ")>0) Or strlsh="" Then
		Call JsAlert("���丨��������Ϣ����","atast_assign.asp")
	End If
	strSql="select * from [mtask] where lsh='"&strlsh&"'"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("ָ����ˮ�ŵ������鲻����,���ʵ!","atask_assign.asp")
	End If
	Rs.Close

	Dim mtbl, tsdzf, tszf, tsxxzlzf, ZstrSql		'ģͷ����, ���Ե��ܷ�, �����ܷ�, ������Ϣ�����ܷ�
	ZstrSql="select * from [mtask] where lsh='"&strlsh&"'"
	Set Rs=xjweb.Exec(ZstrSql, 1)
		mtbl=Rs("mtbl")
		tsdzf=Rs("tsdzf")
		tszf=Rs("tszf")
		tsxxzlzf=Rs("tsxxzlzf")

	strSql=""
	Select case strfplr
		case "��ʼģͷ���Ե�"
			strSql="update [mtask] set mttsdr='"&strzrr&"', mttsdks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "��ʼ���͵��Ե�"
			strSql="update [mtask] set dxtsdr='"&strzrr&"', dxtsdks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "��ʼȫ�׵��Ե�"
			strSql="update [mtask] set mttsdr='"&strzrr&"', mttsdks='"&now()&"', dxtsdr='"&strzrr&"', dxtsdks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)

		case "��ʼģͷ����"
			strSql="update [mtask] set mttsr='"&strzrr&"', mttsks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
			call sendmsg(Rs("mtjgr"), web_info(0), "ģͷ���Կ�ʼ", "��ˮ�� <b>"&strlsh&"</b> ��ʼģͷ����</a>")

			If xjweb.RsCount("[ts_mould] where lsh='"&strlsh&"'")=0 Then
				strSql="insert into ts_mould (lsh,tskssj) values ('"&strlsh&"','"&now()&"')"
				call xjweb.Exec(strSql, 0)
			End If
		case "��ʼ���͵���"
			strSql="update [mtask] set dxtsr='"&strzrr&"', dxtsks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
			call sendmsg(Rs("dxjgr"), web_info(0), "���͵��Կ�ʼ", "��ˮ�� <b>"&strlsh&"</b> ��ʼ���͵���</a>")

			If xjweb.RsCount("[ts_mould] where lsh='"&strlsh&"'")=0 Then
				strSql="insert into ts_mould (lsh,tskssj) values ('"&strlsh&"','"&now()&"')"
				call xjweb.Exec(strSql, 0)
			End If

		case "��ʼȫ�׵���"
			strSql="update [mtask] set mttsr='"&strzrr&"', mttsks='"&now()&"', dxtsr='"&strzrr&"', dxtsks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
			call sendmsg(Rs("mtjgr"), web_info(0), "ȫ�׵��Կ�ʼ", "��ˮ�� <b>"&strlsh&"</b> ��ʼȫ�׵���</a>")
			call sendmsg(Rs("dxjgr"), web_info(0), "ȫ�׵��Կ�ʼ", "��ˮ�� <b>"&strlsh&"</b> ��ʼȫ�׵���</a>")

			strSql="insert into ts_mould (lsh,tskssj) values ('"&strlsh&"','"&now()&"')"
			call xjweb.Exec(strSql, 0)

		case "��ʼģͷ������Ϣ����"
			strSql="update [mtask] set mttsxxzlr='"&strzrr&"', mttsxxzlks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "��ʼ���͵�����Ϣ����"
			strSql="update [mtask] set dxtsxxzlr='"&strzrr&"', dxtsxxzlks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "��ʼȫ�׵�����Ϣ����"
			strSql="update [mtask] set mttsxxzlr='"&strzrr&"', mttsxxzlks='"&now()&"', dxtsxxzlr='"&strzrr&"', dxtsxxzlks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)

		case "����ģͷ���Ե�"
			if not bataskend("mttsdjs", strlsh) then
				strSql="update [mtask] set mttsdjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "mttsd")
			end if

		case "�������͵��Ե�"
			if not bataskend("dxtsdjs", strlsh) then
				strSql="update [mtask] set dxtsdjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "dxtsd")
			end if

		case "����ȫ�׵��Ե�"
			if not bataskend("mttsdjs", strlsh) and not bataskend("dxtsdjs", strlsh) then
				strSql="update [mtask] set mttsdjs='"&now()&"', dxtsdjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "qttsd")
			end if

		case "����ģͷ����"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "mtts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ģͷ���Խ���", "��ˮ�� <b>"&strlsh&"</b> ����ģͷ����</a>")
			end if
		case "�������͵���"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "dxts")
				call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "���͵��Խ���", "��ˮ�� <b>"&strlsh&"</b> �������͵���</a>")
			end if
		case "����ȫ�׵���"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "qtts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ȫ�׵��Խ���", "��ˮ�� <b>"&strlsh&"</b> ����ȫ�׵���</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "ȫ�׵��Խ���", "��ˮ�� <b>"&strlsh&"</b> ����ȫ�׵���</a>")
			end if
		case "ģͷ���ڳ���"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "mtcts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ģͷ���ڳ���", "��ˮ�� <b>"&strlsh&"</b> ģͷ���ڳ���</a>")
			end if
		case "���ͳ��ڳ���"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "dxcts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "���ͳ��ڳ���", "��ˮ�� <b>"&strlsh&"</b> ���ͳ��ڳ���</a>")
			end if
		case "ȫ�׳��ڳ���"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "qtcts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ȫ�׳��ڳ���", "��ˮ�� <b>"&strlsh&"</b> ȫ�׳��ڳ���</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "ȫ�׳��ڳ���", "��ˮ�� <b>"&strlsh&"</b> ȫ�׳��ڳ���</a>")
			end if
		case "ģͷ���⾫��"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "mtjts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ģͷ���⾫��", "��ˮ�� <b>"&strlsh&"</b> ģͷ���⾫��</a>")
			end if
		case "���ͳ��⾫��"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "dxjts")
				call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "���ͳ��⾫��", "��ˮ�� <b>"&strlsh&"</b> ���ͳ��⾫��</a>")
			end if
		case "ȫ�׳��⾫��"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "qtjts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ȫ�׳��⾫��", "��ˮ�� <b>"&strlsh&"</b> ȫ�׳��⾫��</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "ȫ�׳��⾫��", "��ˮ�� <b>"&strlsh&"</b> ȫ�׳��⾫��</a>")
			end if
		case "ģͷԤ���ջ����"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "mtjyts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ģͷԤ���ջ����", "��ˮ�� <b>"&strlsh&"</b> ģͷԤ���ջ����</a>")
			end if
		case "����Ԥ���ջ����"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "dxjyts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "����Ԥ���ջ����", "��ˮ�� <b>"&strlsh&"</b> ����Ԥ���ջ����</a>")
			end if
		case "ȫ��Ԥ���ջ����"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "qtjyts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ȫ��Ԥ���ջ����", "��ˮ�� <b>"&strlsh&"</b> ȫ��Ԥ���ջ����</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "ȫ��Ԥ���ջ����", "��ˮ�� <b>"&strlsh&"</b> ȫ��Ԥ���ջ����</a>")
			end if
		case "ģͷ��������"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "mtysts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ģͷ��������", "��ˮ�� <b>"&strlsh&"</b> ģͷ��������</a>")
			end if
		case "������������"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "dxysts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "������������", "��ˮ�� <b>"&strlsh&"</b> ������������</a>")
			end if
		case "ȫ����������"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'ȷ��ģ�ߵ����Ƿ����
				call fentodb(strlsh, "qtysts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "ȫ����������", "��ˮ�� <b>"&strlsh&"</b> ȫ����������</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "ȫ����������", "��ˮ�� <b>"&strlsh&"</b> ȫ����������</a>")
			end if
		case "����ģͷ������Ϣ����"
			if not bataskend("mttsxxzljs", strlsh) then
				strSql="update [mtask] set mttsxxzljs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "mttsxxzl")
			end if
		case "�������͵�����Ϣ����"
			if not bataskend("dxtsxxzljs", strlsh) then
				strSql="update [mtask] set dxtsxxzljs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "dxtsxxzl")
			end if
		case "����ȫ�׵�����Ϣ����"
			if not bataskend("mttsxxzljs", strlsh) and not bataskend("dxtsxxzljs", strlsh) then
				strSql="update [mtask] set mttsxxzljs='"&now()&"', dxtsxxzljs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				'call fentodb(strlsh, "qttsxxzl")
				call fentodb(strlsh, "mttsxxzl")
				call fentodb(strlsh, "dxtsxxzl")
			end if

		case else
			Call JsAlert("(ϵͳ�쳣)����ϵ����Ա!","")
	end select
	Rs.Close

	'�ж��Ƿ�ȫ�����! �Լ��Ƿ���Ե���
	dim bok
	bok=false
	strSql="select * from [mtask] where lsh='"&strlsh&"'"
	set rs=xjweb.Exec(strSql, 1)
	select case rs("mjxx")
		case "ȫ��"
			if not(isnull(rs("mttsxxzljs"))) and not(isnull(rs("dxtsxxzljs"))) then
				strSql="update [mtask] set mjjs=true where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				bok=true
			end if
		case "ģͷ"
			if not(isnull(rs("mttsxxzljs"))) then
				strSql="update [mtask] set mjjs=true where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				bok=true
			end if
		case "����"
			if not(isnull(rs("dxtsxxzljs"))) then
				strSql="update [mtask] set mjjs=true where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				bok=true
			end if
	end select
	rs.close

	If bok Then
		Call JsAlert("��ˮ�� ��"&strlsh&"�� ����ȫ�����!","atask_assign.asp")
	Else
		Call JsAlert("��ˮ�� ��"&strlsh&"�� ����������� ��"&strfplr&"�� �ɹ�!", "atask_assign.asp?s_lsh="&strlsh&"")

	End If


Rem ����
function bataskend(trs, lsh)			'�ж������Ƿ��Ѿ����,��ֹ�������μǷ�
	bataskend=true
	strSql="select "&trs&" from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	if isnull(rs(trs)) then bataskend=false
	rs.close
	if bataskend then
		response.write("<script language=""javascript"">alert('�������Ѿ�����!');location.href='atask_assign.asp';</script>")
		response.end
	end if
end function

function fentodb(lsh, strlr)
	'����ֵд���ֵ��
	dim itsxs, itslb, iedsx, iedxx, itscs, itsdfz, itsllfz, itsfz, itsxxzlfz, sngmtbl, strzrr, dtjssj, sjjssj, jhsj
	dim ijt, ifgbl, ifcbl, irwnr
	ijt=0

	strSql="select * from [c_fzbl]"
	set rs=xjweb.Exec(strSql, 1)
	ifgbl=CSng(rs("fgbl"))
	ifcbl=CSng(rs("fcbl"))
	rs.close

	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql, 1)
	irwnr=rs("rwlr")
	jhsj=rs("jhjssj")
	itsdfz=rs("tsdzf")
	itslb=rs("tslb")
	itsllfz=rs("tszf")
	itsxxzlfz=rs("tsxxzlzf")
	sngmtbl=rs("mtbl") / 100
	Rs.Close

	select case irwnr
		case "���"
			irwnr=""
			ifgbl=1
		case "����"
			ifgbl=ifcbl
	end select

	'���⾫��
	If Right(strlr,3)="jts" Then
		strSql="select * from [ts_tsxx] where lsh='"&lsh&"' and tslr like '%����%'"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		If Rs.Eof Or Rs.Bof Then
			rs.addnew
				rs("lsh")=lsh
				rs("tsyy")="���⾫��"
				rs("tslr")="���⾫��"
				rs("tssj")=now()
				rs("tsr")="Sj901"
			rs.update
			ijt=1
		End If
		rs.close

		If ijt=1 Then
			strSql="select * from [ts_mould] where lsh='"&lsh&"'"
			Call xjweb.Exec("",-1)
			rs.open strSql,conn,1,3
				if isnull(rs("tskssj")) then rs("tskssj")=now()
				rs("tscs")=rs("tscs") + 1
				rs("tsgxsj")=now()
			rs.update
			rs.close
		End If
	End If
	'Ԥ���ջ�������������ա����ڳ���
	If InStr(strlr,"jyts")>0 Then itsllfz=itsllfz*1.5
	If InStr(strlr,"ysts")>0 Then itsllfz=itsllfz*2
	If InStr(strlr,"cts")>0 Then itsllfz=itsllfz*0.75
	'================
	itsxs=1
	If not(IsNull(itslb)) and itslb<>"A��" and right(strlr,2)="ts" and right(strlr,3)<>"cts" Then
		strSql="select * from [c_tscs] where dmlb='"&itslb&"'"
		set rs=xjweb.Exec(strSql, 1)
			If not(Rs.Eof Or Rs.Bof) Then
				iedsx=rs("edsx")
				iedxx=rs("edxx")
			else
				iedsx=0
			End If
		Rs.Close
		If iedsx<>0 Then
			strSql="select * from [ts_mould] where LSH='"&lsh&"'"
			set rs=xjweb.Exec(strSql, 1)
			itscs=rs("TSCS")
			Rs.Close
			If iedxx > itscs Then itsxs=1+(iedxx-itscs)*0.15
			If iedsx < itscs Then itsxs=1+(iedsx-itscs)*0.15
			If itsxs < 0.6 Then itsxs=0.6
		End If
	End If
	'Ԥ���ջ��������������ֻ�ӷֲ�����
	If InStr(strlr,"jyts")>0 or InStr(strlr,"ysts")>0 Then
		If itsxs<1 Then
			itsxs=1
		End If
	End If
	'ʵ�ʵ��Է�ֵ
	itsfz=Round(itsllfz*itsxs,1)

	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	sjjssj=rs("sjjssj")
	select case strlr
		case "mttsd"	'ģͷ���Ե�
			strzrr=rs("mttsdr")
			dtjssj=rs("mttsdjs")
'			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','ģͷ���Ե�',"&itsdfz*sngmtbl*ifgbl&",'"&dtjssj&"',0,'"&strzrr&"')"
'			call xjweb.Exec(strSql,0)
			If (datediff("d", dtjssj, jhsj) < -20) and (not isnull(rs("mttsdjs"))) and (not isnull(rs("dxtsdjs"))) then
				call Tsdkp(strzrr, strlsh, 5, strlsh)
			End if
		case "dxtsd"	'���͵��Ե�
			strzrr=rs("dxtsdr")
			dtjssj=rs("dxtsdjs")
'			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','���͵��Ե�',"&itsdfz*(1-sngmtbl)*ifgbl&",'"&dtjssj&"',0,'"&strzrr&"')"
'			call xjweb.Exec(strSql,0)
			If (datediff("d", dtjssj, jhsj) < -20) and (not isnull(rs("mttsdjs"))) and (not isnull(rs("dxtsdjs"))) then
				call Tsdkp(strzrr, strlsh, 5, strlsh)
			End if
		case "qttsd"	'ȫ�׵��Ե�
			strzrr=rs("mttsdr")
			dtjssj=rs("mttsdjs")
'			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','ȫ�׵��Ե�',"&itsdfz*ifgbl&",'"&dtjssj&"',0,'"&strzrr&"')"
'			call xjweb.Exec(strSql,0)
			If (datediff("d", dtjssj, jhsj) < -20) and (not isnull(rs("mttsdjs"))) and (not isnull(rs("dxtsdjs"))) then
				call Tsdkp(strzrr, strlsh, 5, strlsh)
			End if
		case "mtts"		'ģͷ����
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ģͷ���Ժϸ�',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���Ժϸ�' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxts"		'���͵���
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','���͵��Ժϸ�',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���Ժϸ�' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtts"		'ȫ�׵���
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ȫ�׵��Ժϸ�',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���Ժϸ�' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtcts"		'ģͷ���ڳ���
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ģͷ���ڳ���',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���ڳ���' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxcts"		'���ͳ��ڳ���
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','���ͳ��ڳ���',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���ڳ���' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtcts"		'ȫ�׳��ڳ���
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ȫ�׳��ڳ���',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���ڳ���' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtjts"		'ģͷ���⾫��
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ģͷ���⾫��',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���⾫��' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxjts"		'���ͳ��⾫��
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','���ͳ��⾫��',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���⾫��' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtjts"		'ȫ�׳��⾫��
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ȫ�׳��⾫��',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='���⾫��' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtjyts"		'ģͷԤ���ջ����
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ģͷԤ���ջ����',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='Ԥ���ջ����' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxjyts"		'����Ԥ���ջ����
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','����Ԥ���ջ����',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='Ԥ���ջ����' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtjyts"		'ȫ��Ԥ���ջ����
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ȫ��Ԥ���ջ����',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='Ԥ���ջ����' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtysts"		'ģͷ��������
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ģͷ��������',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='��������' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxysts"		'������������
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','������������',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='��������' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtysts"		'ȫ����������
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','ȫ����������',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='��������' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mttsxxzl"		'ģͷ������Ϣ����
			strzrr=rs("mttsxxzlr")
			dtjssj=rs("mttsxxzljs")
			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','ģͷ������Ϣ����',"&itsxxzlfz*sngmtbl&",'"&dtjssj&"',0,'"&strzrr&"')"
			call xjweb.Exec(strSql,0)
		case "dxtsxxzl"		'���͵�����Ϣ����
			strzrr=rs("dxtsxxzlr")
			dtjssj=rs("dxtsxxzljs")
			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','���͵�����Ϣ����',"&itsxxzlfz*(1-sngmtbl)&",'"&dtjssj&"',0,'"&strzrr&"')"
			call xjweb.Exec(strSql,0)
		case "qttsxxzl"		'ȫ�׵�����Ϣ����
			strzrr=rs("mttsxxzlr")
			dtjssj=rs("mttsxxzljs")
			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','ȫ�׵�����Ϣ����',"&itsxxzlfz&",'"&dtjssj&"',0,'"&strzrr&"')"
			call xjweb.Exec(strSql,0)
		case else
	end select
end function

function check_tsend(strlsh)
	strSql="select * from [mtask] where lsh='"&strlsh&"'"
	set rs=xjweb.Exec(strSql, 1)
	if not(rs.eof or rs.bof) then
		select case rs("mjxx")
			case "ȫ��"
				if not(isnull(rs("mttsjs"))) and not(isnull(rs("dxtsjs"))) then
					strSql="update [ts_mould] set tsjssj='"&now()&"' where lsh='"&strlsh&"'"
					call xjweb.Exec(strSql, 0)
				end if
			case "ģͷ"
				if not(isnull(rs("mttsjs"))) then
					strSql="update [ts_mould] set tsjssj='"&now()&"' where lsh='"&strlsh&"'"
					call xjweb.Exec(strSql, 0)
				end if
			case "����"
				if not(isnull(rs("dxtsjs"))) then
					strSql="update [ts_mould] set tsjssj='"&now()&"' where lsh='"&strlsh&"'"
					call xjweb.Exec(strSql, 0)
				end if
		end select
	end if
	rs.close
end function

'Function Tsdkp(lsh)	'�����Ե�����ʱȡ��ģ�ߵ���ǰ��,�ݲ�ִ��
'	Dim tmpRs
'	tmpRs="delete from [kp_jsb] where kp_lsh='"&lsh&"'and kp_item='��ǰ'"
'	call xjweb.Exec(tmpRs, 0)
'End Function

Function Tsdkp(strzrr, lsh, iprice, bz)
	dim tmpSql, tmpRs, iGroup, iKPKind, strKpTopic, strKpItem, ikpmul, strbz, strSql
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strzrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

	iKpKind=5		'5Ϊ��Ա����
	strKpTopic="�������"
		strKpItem="�ӳ�"
		iKpMul=-1
		strBz=bz
		tmpSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		set tmpRs=Server.CreateObject("adodb.recordset")
		tmpRs.open tmpSql,conn,1,3
		tmpRs.AddNew
			tmpRs("kp_time")=Now()
			tmpRs("kp_zrr")=strZrr
			tmpRs("kp_zrrjs")="�����ֲ�"
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

Function TscsKp(slsh)	'���ݵ��Դ����Լ���Ա���鳤���п���
	Dim KpSql, KpRs, sngmtbl, iddh, idmmc, ijgzz, ijsdb, igfwh, imtjgr, imtjgshr, idxjgr, idxjgshr, imtsjr, imtsjshr, idxsjr, idxsjshr, imjzf, itslb, iedsx, iedxx, itscs, itsxs, ifz, iGroup, strKpItem, iKpMul, kp_bz, iZlID
	'�������������ڷֱ��Ƿ�ͬһ���������
	Randomize
	iZlID=rnd*99999
	'��ȡ����ģ��ʵ�ʵ��Դ���
	KpSql="select * from [ts_mould] where lsh='"&slsh&"'"
	set KpRs=xjweb.Exec(KpSql, 1)
	if not(KpRs.eof or KpRs.bof) then
		If isnull(KpRs("tsjssj")) then
			KpRs.close
			exit Function
		else
			itscs=KpRs("tscs")
		End If
	end if
	KpRs.close
	'��ȡģ�ߵ������͡�����˵���Ϣ
	KpSql="select * from [mtask] where lsh='"&slsh&"'"
	set KpRs=xjweb.Exec(KpSql, 1)
		ijgzz=KpRs("jgzz")						'�鳤
		imtjgr=KpRs("mtjgr")					'ģͷ�ṹ��
		imtjgshr=KpRs("mtjgshr")				'ģͷ�ṹ�����
		idxjgr=KpRs("dxjgr")					'���ͽṹ��
		idxjgshr=KpRs("dxjgshr")				'���ͽṹ�����
		itslb=KpRs("tslb")						'�������
	KpRs.Close
	'��ȡ����ģ�߶���Դ���
	If isnull(itslb) Then exit Function
	KpSql="select * from [c_tscs] where dmlb='"&itslb&"'"
	set KpRs=xjweb.Exec(KpSql, 1)
		If not(KpRs.Eof Or KpRs.Bof) Then
			iedsx=KpRs("edsx")
			iedxx=KpRs("edxx")
		else
			KpRs.Close
			exit Function
		End If
	KpRs.Close
	'���㳬����׼Ӧ���˵��Դ���
	itsxs=0
	If iedxx > itscs Then itsxs=iedxx-itscs
	If iedsx < itscs Then itsxs=iedsx-itscs
	If itsxs=0 or (Instr("B��C��",itslb)>0 and itsxs<0) Then
		 exit Function
	End If
	'=====================
	'�Խṹ����˽��п���
'	If not(isNull(imtjgr)) Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&imtjgr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="���ڵ������ڶ����"
'		else
'			strKpItem="���ڵ��Զ��ڶ����"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=imtjgr
'			KpRs("kp_zrrjs")="ģͷ�ṹ"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="�������"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"�:"&iedxx&"-"&iedsx&"��,ʵ��:"&itscs&"��"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
'
'	If not(isNull(idxjgr)) Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&idxjgr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="���ڵ������ڶ����"
'		else
'			strKpItem="���ڵ��Զ��ڶ����"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=idxjgr
'			KpRs("kp_zrrjs")="���ͽṹ"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="�������"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"�:"&iedxx&"-"&iedsx&"��,ʵ��:"&itscs&"��"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
'	'�Խṹ����˽��п���
'	If not(isNull(imtjgshr)) and imtjgshr<>ijgzz Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&imtjgshr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="���ڵ������ڶ����"
'		else
'			strKpItem="���ڵ��Զ��ڶ����"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=imtjgshr
'			KpRs("kp_zrrjs")="ģͷ�ṹ���"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="�������"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"�:"&iedxx&"-"&iedsx&"��,ʵ��:"&itscs&"��"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
'	If not(isNull(idxjgshr)) and idxjgshr<>ijgzz Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&idxjgshr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="���ڵ������ڶ����"
'		else
'			strKpItem="���ڵ��Զ��ڶ����"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=idxjgshr
'			KpRs("kp_zrrjs")="���ͽṹ���"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="�������"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"�:"&iedxx&"-"&iedsx&"��,ʵ��:"&itscs&"��"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
	'���鳤���п���(�ṹ�鳤(ijgzz)�������鳤(tszz))
'	Dim xxzz, tszz
'	KpSql="Select * from [ims_user] where mid(user_able,4,1)>0"
'	Set KpRs=xjweb.Exec(KpSql,1)
'		Do While Not KpRs.Eof
'			If KpRs("user_group")=5 and mid(KpRs("user_able"),6,1)>0 and mid(KpRs("user_able"),3,1)=0 Then tszz=KpRs("user_name")
'			KpRs.moveNext
'		Loop
'	KpRs.Close

'	If not(isNull(ijgzz)) Then
'		ifz=Round(imjzf*sngmtbl*0.1*itsxs,1)
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&ijgzz&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="���ڵ������ڶ����"
'		else
'			strKpItem="���ڵ��Զ��ڶ����"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=ijgzz
'			KpRs("kp_zrrjs")="�ṹ�鳤"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=3
'			KpRs("kp_topic")="�������"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs*0.4
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"�:"&iedxx&"-"&iedsx&"��,ʵ��:"&itscs&"��"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If

'	KpSql="Select [user_group] from [ims_user] where [user_name]='"&tszz&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'	KpRs.Close
'	KpSql="select * from [kp_jsb]"
'	Call xjweb.Exec("",-1)
'	KpRs.open KpSql,conn,1,3
'	KpRs.AddNew
'		KpRs("kp_time")=Now()
'		KpRs("kp_zrr")=tszz
'		KpRs("kp_zrrjs")="�����鳤"
'		KpRs("kp_group")=iGroup
'		KpRs("kp_kind")=3
'		KpRs("kp_topic")="�����֤"
'		KpRs("kp_item")=strKpItem
'		KpRs("kp_uprice")=itsxs*0.2
'		KpRs("kp_mul")=iKpMul
'		KpRs("kp_bz")=slsh&"�:"&iedxx&"-"&iedsx&"��,ʵ��:"&itscs&"��"
'		KpRs("kp_kpr")="Sj901"
'		KpRs("kp_lsh")=slsh
'		KpRs("kp_zlid")=iZlID
'		KpRs.Update
'	KpRs.Close

	set KpRs = nothing
End Function
%>