<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
dim action, iid
action=request("action")
iid=request("id")
select case action
	case "affirm"		'��֤�յ�
		strSql="update [ims_message] set delr=1 where id="&clng(iid)&""
		Call xjweb.Exec(strSql, 0)
		response.write("<script language=""javascript"">self.close();</script>")
	case "delete"		'ɾ������
		dim item
		if request("boxkind")="incept" then
			for each item in request.form("chkbox")
				strSql="update [ims_message] set delr=2 where id="&item&""
				Call xjweb.Exec(strSql, 0)
			next
			Call JsAlert("����ɾ���ɹ�!","uctrl_dismsg.asp?box=incept")
		else
			for each item in request.form("chkbox")
				strSql="update [ims_message] set dels=2 where id="&item&""
				Call xjweb.Exec(strSql, 0)
			next
			Call JsAlert("����ɾ���ɹ�!","uctrl_dismsg.asp?box=send")
		end if
	case "Sdelete"		'��֤�յ���ɾ��
		strSql="update [ims_message] set delr=2 where id="&clng(iid)&""
		Call xjweb.Exec(strSql, 0)
		response.write("<script language=""javascript"">self.close();</script>")
	case "send"			'���Ͷ���
		dim struser, strtitle, strcontent
		struser=request("incept")
		strtitle=trim(request("title"))
		strcontent=request("content")
		if struser="" or strtitle="" or strcontent="" then
			Call JsAlert("��ȷ�ϴ���ȷ��ڽ��벢��֤��Ϣ��������!","")
		else
			struser=split(struser,"|")
			for i=0 to ubound(struser)
				strSql="insert into [ims_message] (sender, incept, title, content) values ('"&session("userName")&"','"&struser(i)&"','"&strtitle&"','"&strcontent&"')"
				Call xjweb.Exec(strSql, 0)
			next
			Call JsAlert("���ŷ��ͳɹ�!",Request.ServerVariables("HTTP_REFERER"))
		end if
	case else
		Call JsAlert(action & "����ϵ����Ա","index.asp")
end select
%>