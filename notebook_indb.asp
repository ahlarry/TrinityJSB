<!--#include file="include/conn.asp"-->
<%
	dim strusername, strcontent, iid, strreply, strininf, strtemp
	if Session("userName") <> "" and Session("userName") <> "客人" then strusername = Session("userName")
	if strusername = "" then strusername = request.servervariables("local_addr")

	strininf = request("indbinf")
	select case strininf
		case "add"
			strcontent = request("lylr")
			strSql="insert into [notebook] (content,username,indate) values ('"&strcontent&"','"&strusername&"','"&now&"')"
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("签写留言成功!","notebook.asp")
		case "change"
			strcontent = request("lylr")
			iid = clng(request("id"))
			strSql="update [notebook] set content = '"&strcontent&"',editdate='"&now&"' where id = "&iid&""
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("更改留言成功!","notebook.asp")
		case "reply"
			strreply = request("hf")
			iid = clng(request("id"))
			strSql="update notebook set reply = '"&strreply&"' where id = "&iid&""
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("留言回复成功!","notebook.asp")
		case "delete"
			iid = clng(request("id"))
			strSql="delete from notebook where id = "&iid&""
			Call xjweb.Exec(strSql, 0)
			Call JsAlert("删除留言成功!","notebook.asp")
		case else
			Call JsAlert("出错提示: " & request("indbinf") & " 。\n请记下上面的内容并联系管理员!","notebook.asp")
	end select
%>