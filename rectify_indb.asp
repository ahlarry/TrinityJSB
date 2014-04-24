<!--#include file="include/conn.asp"-->
<%
	Call ChkAble(11)
	Dim action, strzrbm, strbh, strxxbm, strjssj, strbhgnr, stryfcsyq, strqxsj, strps,  stryyfx, strjzcs, strlsqk, stryzjl, strwczk, iid
	action="" : strzrbm="" : strbh="" : strxxbm="" : strjssj="" : strbhgnr="" : stryfcsyq="" : strqxsj="" : strps="" : stryyfx="" : strjzcs="" : strlsqk="" : stryzjl="" : strwczk="纠正中" : iid=0
	action=LCase(Request("action"))
	strzrbm=Request("zrbm")
	strbh=Request("bh")
	strxxbm=Request("xxbm")
	strjssj=Request("jssj")
	strbhgnr=Request("bhgnr")
	stryfcsyq=Request("yfcsyq")
	strqxsj=Request("qxsj")
	strps=Request("ps")
	stryyfx=Request("yyfx")
	strjzcs=Request("jzcs")
	strlsqk=Request("lsqk")
	stryzjl=Request("yzjl")
	If IsNumeric(Trim(Request("id"))) Then iid=CLng(Trim(Request("id")))

	'数据入库函数从这里开始
	Select Case action
		Case "add"
			If strzrbm="" Or strjssj="" Or strbhgnr="" Then
				Call JsAlert("请确认信息输入完整!请从正确的入口进入!","Rectify_list.asp")
			Else
				Call Rectify_Add()
			End If
		Case "change"
			If strzrbm="" Or strbhgnr="" Or iid="" Then
				Response.Write strbhgnr
				Call JsAlert("请确认信息输入完整!","")
			Else
				Call Rectify_Change()
			End If
		Case "delete"
			If iid=0 Then
				Call JsAlert("请确认从系统入口进入!","")
			Else
				strSql="delete from [rectify] where id=" & iid
				Call xjweb.Exec(strSql, 0)
				Call JsAlert("外部质量信息删除成功","Rectify_list.asp")
			end if
		Case Else
			Call JsAlert("action="&action&", 请联系管理员!","Rectify_list.asp")
	End Select
	
	'验证外部信息完成情况
	Function Rectify_zt()
		If stryfcsyq <> "" Then strwczk = "分析中"		 
		If stryyfx <> "" Then strwczk = "制定措施中" 
		If strjzcs <> "" Then strwczk = "跟踪检查中" 
		If strlsqk <> "" Then strwczk = "验证中"  
		If stryzjl <> "" Then strwczk = "已闭环"  
	End Function
	
	'外部质量信息 入库
	Function Rectify_Add()
		strSql="select * from [Rectify]"
		Call xjweb.Exec("",-1)
		Call Rectify_zt()
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("zrbm")=strzrbm
			Rs("bh")=strbh
			Rs("xxbm")=strxxbm
			Rs("jssj")=strjssj
			Rs("bhgnr")=strbhgnr
			Rs("yfcsyq")=stryfcsyq
			Rs("qxsj")=strqxsj
			Rs("ps")=strps
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl
			Rs("wczk")=strwczk
		Rs.Update
		Rs.Close
		Call JsAlert("纠正/预防措施表添加成功","Rectify_add.asp")
	End Function

	'更改外部质量信息 入库	
	Function Rectify_Change()
		'检测ID号是否存在
		Set Rs=xjweb.Exec("select * from [Rectify] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("编号 " & iid & " 纠正/预防措施表不存在！","Rectify_list.asp")
			Rs.Close
			Exit Function
		End If
		Rs.Close

		strSql="select * from [Rectify] where id=" & iid
		Call xjweb.Exec("",-1)
		Call Rectify_zt()
		'strmsg="数据库操作"
		Rs.open strSql,conn,1,3
			Rs("zrbm")=strzrbm
			Rs("bh")=strbh
			Rs("xxbm")=strxxbm
'			Rs("jssj")=strjssj
			Rs("bhgnr")=strbhgnr
			Rs("yfcsyq")=stryfcsyq
			Rs("qxsj")=strqxsj
			Rs("ps")=strps
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl
			Rs("wczk")=strwczk
		Rs.update
		Rs.close

		Call JsAlert("纠正/预防措施表更改成功","Rectify_list.asp")
	End Function
%>