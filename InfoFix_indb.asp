<!--#include file="include/conn.asp"-->
<%
'8:53 2007-3-12-����һ
	'���ļ��������������Ϣ������������
	Dim action
	action=Request("action")
	Dim strlsh, strzxr, strxs, strjhjssj, strNum, strwc
	'������ʼ��
	strzxr=""
	strNum=Trim(Request("Num"))
	strjhjssj=Request("psy") & "��" & Request("psm") & "��" & Request("psd") & "��"

	'������⺯�������￪ʼ
	Select Case action
		Case "add"
			Call InfoFix_add()
		Case "change"
			Call InfoFix_change()
		Case else
			response.write "action=" & action
	End select

	'ִ�����������
	Function InfoFix_add()
		dim tmp2Sql, tmp2Rs
		For i=1 to strNum
		strlsh=Trim(Request("lsh"&i)) : strzxr=Trim(Request("zxr"&i))
		strxs=Trim(Request("fzxs"&i)) : strwc=Trim(Request("xtzlwc"&i))
			If strzxr<>"" and strwc="�������" and not bmtaskend("xtxxsjjs", strlsh) Then
				strSql="update [mtask] set xtxxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "������Ϣ����")
				call FenToDB(strlsh)
			End If
			If strzxr<>"" and strxs<>"" Then
				strSql="update [mtask] set xtxxzlr='"&strzxr&"', xtxxzlxs='"&strxs&"', xtxxzlks='"&now()&"', xtxxjhjs='"&strjhjssj&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "������Ϣ����", strzxr, now())
				call sendmsg(strzxr, web_info(0), "������Ϣ����", "��ˮ�� <b>"&strlsh&"</b> ������Ϣ����</a>")
			End If
		Next
		Call JsAlert("��������ӳɹ�!", "InfoFix_add.asp")
		Response.End
	End Function

	'�������������
	Function InfoFix_change()
		dim tmp2Sql, tmp2Rs
		For i=1 to strNum
		strlsh=Trim(Request("lsh"&i)) : strzxr=Trim(Request("zxr"&i))
			If strzxr<>"" and not bmtaskend("xtxxsjjs", strlsh) Then
				strSql="update [mtask] set xtxxzlr='"&strzxr&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
			End If
		Next
		Call JsAlert("��������ĳɹ�!", "InfoFix_zzchange.asp")
		Response.End
	End Function

function taskstart(lsh, rwlr, zrr, kssj)		'����ʼ
	strSql="insert into [mtask_cur] (lsh, rwlr, zrr, kssj) values ('"&lsh&"', '"&rwlr&"', '"&zrr&"','"&kssj&"')"
	call xjweb.Exec(strSql, 0)
end function

function taskend(lsh, rwlr)						'�������
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
	Dim mjfz, fzxs, ijc, ijc2
	'ijc===���ͷ�ֵ
	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	mjfz=rs("mjzf")
	fzxs=Rs("xtxxzlxs")
	ijc2=datediff("d",rs("xtxxsjjs"),rs("xtxxjhjs"))
	ijc=0
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc,zrr) values ('"&rs("lsh")&"','������Ϣ����',"&Round(mjfz*fzxs,1)&",'"&rs("xtxxsjjs")&"',"&ijc&",'"&rs("xtxxzlr")&"')"
				call xjweb.Exec(strSql,0)
				If ijc2<0 Then Call ygkptodb(rs("xtxxzlr"), ijc2, 1, rs("lsh"), rs("lsh")&"������Ϣ����")
	rs.close
end function

Function ygkptodb(zrr, tt, iprice, strlsh, lsh)
	dim tmpSql, tmpRs, iGroup, iKPKind, strKpTopic, strKpItem, ikpmul, strbz, strZrr
	strZrr=zrr
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

	iKpKind=5		'5Ϊ��Ա����
	strKpTopic="�������"
	If tt>0 Then
		strKpItem="��ǰ"
	Else
		strKpItem="����ӳ�"
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
		tmpRs("kp_time")=Now()
		tmpRs("kp_zrr")=strZrr
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
%>
