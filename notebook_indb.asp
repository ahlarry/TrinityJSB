<!--#include file="include/conn.asp"-->
<%
	dim strusername, strcontent, iid, strreply, strininf, strtemp
	if Session("userName") <> "" and Session("userName") <> "����" then strusername = Session("userName")
	if strusername = "" then strusername = request.servervariables("local_addr")

	strininf = request("indbinf")
	select case strininf
		case "add"
			strcontent = request("lylr")
			strSql="insert into [notebook] (content,username,indate) values ('"&strcontent&"','"&strusername&"','"&now&"')"
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("ǩд���Գɹ�!","notebook.asp")
		case "change"
			strcontent = request("lylr")
			iid = clng(request("id"))
			strSql="update [notebook] set content = '"&strcontent&"',editdate='"&now&"' where id = "&iid&""
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("�������Գɹ�!","notebook.asp")
		case "reply"
			strreply = request("hf")
			iid = clng(request("id"))
			strSql="update notebook set reply = '"&strreply&"' where id = "&iid&""
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("���Իظ��ɹ�!","notebook.asp")
		case "delete"
			iid = clng(request("id"))
			strSql="delete from notebook where id = "&iid&""
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("ɾ�����Գɹ�!","notebook.asp")
		case else
			Call JsAlert("������ʾ: " & request("indbinf") & " ��\n�������������ݲ���ϵ����Ա!","notebook.asp")
	end select
%>