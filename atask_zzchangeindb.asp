<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble(4)
	'���ļ�ֻ���鳤��������������,������ͬʱ���ķ�ֵ��
	dim strmttsdr, strdxtsdr, strmttsr, strdxtsr, strmttsxxzlr, strdxtsxxzlr, strlsh
	strmttsdr="" : strdxtsdr=""
	strmttsr="" : strdxtsr=""
	strmttsxxzlr="" : strdxtsxxzlr=""
	strlsh=""

	strmttsdr=request("mttsdr")
	strmttsr=request("mttsr")
	strmttsxxzlr=request("mttsxxzlr")
	strdxtsdr=request("dxtsdr")
	strdxtsr=request("dxtsr")
	strdxtsxxzlr=request("dxtsxxzlr")
	strlsh=request("lsh")

	if strlsh="" then
		Call JsAlert("��ˮ�Ų�ȷ��,�޷�ȷ������ĵ�������!\n���������ڽ���!","atask_zzchange.asp")
	end if

	set rs=xjweb.Exec("select lsh from [mtask] where lsh='"&strlsh&"'",1)
	if rs.eof or rs.bof then
		Call JaAlert("��ˮ�� ��"&strlsh&"�� ���񲻴���!���ʵ!","atask_zzchange.asp")
	end if
	Rs.close

	strSql="select * from [mtask] where lsh='" & strlsh & "'"
	call xjweb.Exec("",-1)
	strmsg=""
	rs.open strSql,conn,1,3
		if strmttsdr<>"" and strmttsdr<>rs("mttsdr") then
			rs("mttsdr")=strmttsdr
			strmsg = strmsg & "����ģͷ���Ե���"
			strSql="update [mantime] set zrr='"&strmttsdr&"' where lsh='"&strlsh&"' and rwlr='ģͷ���Ե�'"
			call xjweb.Exec(strSql, 0)
		end if
		if strdxtsdr<>"" and strdxtsdr<>rs("dxtsdr") then 
			rs("dxtsdr")=strdxtsdr
			strmsg = strmsg & "���Ķ��͵��Ե���"
			strSql="update [mantime] set zrr='"&strdxtsdr&"' where lsh='"&strlsh&"' and rwlr='���͵��Ե�'"
			call xjweb.Exec(strSql, 0)
		end if
		if strmttsr<>"" and strmttsr<>rs("mttsr") then 
			rs("mttsr")=strmttsr
			strmsg = strmsg & "����ģͷ������"
			strSql="update [mantime] set zrr='"&strmttsr&"' where lsh='"&strlsh&"' and rwlr='ģͷ����'"
			call xjweb.Exec(strSql, 0)
		end if
		if strdxtsr<>"" and strdxtsr<>rs("dxtsr") then
			rs("dxtsr")=strdxtsr 
			strmsg = strmsg & "���Ķ��͵�����"
			strSql="update [mantime] set zrr='"&strdxtsr&"' where lsh='"&strlsh&"' and rwlr='���͵���'"
			call xjweb.Exec(strSql, 0)
		end if
		if strmttsxxzlr<>"" and strmttsxxzlr<>rs("mttsxxzlr") then 
			rs("mttsxxzlr")=strmttsxxzlr
			strmsg = strmsg & "����ģͷ������Ϣ������"
			strSql="update [mantime] set zrr='"&strmttsxxzlr&"' where lsh='"&strlsh&"' and rwlr='ģͷ������Ϣ����'"
			call xjweb.Exec(strSql, 0)
		end if
		if strdxtsxxzlr<>"" and strdxtsxxzlr<>rs("dxtsxxzlr") then
			rs("dxtsxxzlr")=strdxtsxxzlr
			strmsg = strmsg & "���Ķ��͵�����Ϣ������"
			strSql="update [mantime] set zrr='"&strdxtsxxzlr&"' where lsh='"&strlsh&"' and rwlr='���͵�����Ϣ����'"
			call xjweb.Exec(strSql, 0)
		end if
	rs.update
	rs.close
	
	If strmsg<>"" Then
		strmsg="���ݿ����:" & strmsg
		strSql="insert into [ims_log] (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','����������','"&strmsg&"','"&now()&"')"
		call xjweb.Exec(strSql,0)
		Call JsAlert("��ˮ�� ��" & strlsh & "�� �������������˸��ĳɹ�!","atask_zzchange.asp")
	Else
		Call JsAlert("��û�ж���ˮ�� ��" & strlsh & "�� ������������κθ���!","atask_zzchange.asp")
	End If
%>