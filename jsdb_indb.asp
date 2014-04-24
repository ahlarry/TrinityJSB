<!--#include file="include/conn.asp"-->
<%
'2013/1/14 10:15
	'添加技术代表任务书入库	
	Dim action, strhth, strkhmc, strdrwnr, strfz, strjhsj, strzz
	action=Request("action")
	strhth=Trim(Request("hth")) : strkhmc=Trim(Request("khmc")) : strdrwnr=Trim(Request("rwnr")) : strjhsj=Trim(Request("jhjssj"))
	strzz=Trim(Request("sjr")) : strfz=NulltoNum(Request("jcf"))

	Select Case action
		Case "add"
			Call jsdb_add()
		Case "change"
			Call jsdb_change()
	End select

	'添加任务书入库
	Function jsdb_add()
	'对待入库数据进行处理
	Dim strMsg
	strMsg=""
	If strhth="" Then strMsg="合同号为空!<br>"
	If strkhmc="" Then strMsg=strMsg & "客户名称为空!<br>"
	If strzz="" Then strMsg=strMsg & "组长为空!<br>"
	If strfz=0 Then strMsg=strMsg & "分值为0!<br>"

	If strMsg <> ""Then
		infoTitle="数据不完整"
		infoContents=strMsg & "<br>点击<a href=""#"" onclick='history.go(-1);'>返回前页</a>重新输入"
		GotoPrompt()
	End If
	
		'检测合同号是否已存在
		strSql="select * from [jsdb] where [hth]='" & strhth & "'"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
			If IsNull(Rs("shjssj")) Then
				Rs("khmc")=strkhmc
				If strdrwnr<>"" Then Rs("rwnr")=strdrwnr
				Rs("zz")=strzz
				Rs("jcf")=strfz
				Rs("jhjssj")=strjhsj
				Rs.update
				Rs.Close
				Call JsAlert("任务更改成功!", "jsdb_add.asp")
			else
				Rs.Close
				Call JsAlert("任务已完成，无法修改!!", "jsdb_add.asp")
			End If
		End If
		Rs.Close
		
		strSql="select * from [jsdb]"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Rs.AddNew
			Rs("hth")=strhth
			Rs("khmc")=strkhmc
			If strdrwnr<>"" Then Rs("rwnr")=strdrwnr
			Rs("zz")=strzz
			Rs("jcf")=strfz
			Rs("jhjssj")=strjhsj
		Rs.update
		Rs.Close
		Call JsAlert("任务书添加成功!", "jsdb_add.asp")
	End Function
%>
