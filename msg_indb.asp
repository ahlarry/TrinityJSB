<!--#include file="include/conn.asp"-->
<%
Call ChkPageAble(0)
dim action, iid
action=request("action")
iid=request("id")
select case action
	case "affirm"		'验证收到
		strSql="update [ims_message] set delr=1 where id="&clng(iid)&""
		Call xjweb.Exec(strSql, 0)
		response.write("<script language=""javascript"">self.close();</script>")
	case "delete"		'删除短信
		dim item
		if request("boxkind")="incept" then
			for each item in request.form("chkbox")
				strSql="update [ims_message] set delr=2 where id="&item&""
				Call xjweb.Exec(strSql, 0)
			next
			Call JsAlert("短信删除成功!","uctrl_dismsg.asp?box=incept")
		else
			for each item in request.form("chkbox")
				strSql="update [ims_message] set dels=2 where id="&item&""
				Call xjweb.Exec(strSql, 0)
			next
			Call JsAlert("短信删除成功!","uctrl_dismsg.asp?box=send")
		end if
	case "Sdelete"		'验证收到并删除
		strSql="update [ims_message] set delr=2 where id="&clng(iid)&""
		Call xjweb.Exec(strSql, 0)
		response.write("<script language=""javascript"">self.close();</script>")
	case "send"			'发送短信
		dim struser, strtitle, strcontent
		struser=request("incept")
		strtitle=trim(request("title"))
		strcontent=request("content")
		if struser="" or strtitle="" or strcontent="" then
			Call JsAlert("请确认从正确入口进入并保证信息输入完整!","")
		else
			struser=split(struser,"|")
			for i=0 to ubound(struser)
				strSql="insert into [ims_message] (sender, incept, title, content) values ('"&session("userName")&"','"&struser(i)&"','"&strtitle&"','"&strcontent&"')"
				Call xjweb.Exec(strSql, 0)
			next
			Call JsAlert("短信发送成功!",Request.ServerVariables("HTTP_REFERER"))
		end if
	case else
		Call JsAlert(action & "请联系管理员","index.asp")
end select
%>