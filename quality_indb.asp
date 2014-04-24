<!--#include file="include/conn.asp"-->
<%
	Call ChkAble(11)
	Dim action, strkhmc, strlxr, strlxdh, strhth, strgzlh, strjssj, strZrr, strzywt, stryjcs, stryyfx, strjzcs, strlsqk, stryzjl, strwczk, iid
	action="" : strkhmc="" : strlxr="" : strlxdh="" : strhth="" : strgzlh="" : strjssj="" : strZrr="" : strzywt="" : stryjcs="" : stryyfx="" : strjzcs="" : strlsqk="" : stryzjl="": strwczk="纠正中" : iid=0
	action=LCase(Request("action"))
	strkhmc=Request("khmc")
	strlxr=Request("lxr")
	strlxdh=Request("lxdh")
	strhth=Request("hth")
	strgzlh=Request("gzlh")
	strjssj=Request("jssj")
	strZrr=Request("zrr")
	strzywt=Request("zywt")
	stryjcs=Request("yjcs")
	stryyfx=Request("yyfx")
	strjzcs=Request("jzcs")
	strlsqk=Request("lsqk")
	stryzjl=Request("yzjl")
	If IsNumeric(Trim(Request("id"))) Then iid=CLng(Trim(Request("id")))

	'数据入库函数从这里开始
	Select Case action
		Case "add"
			If strhth="" Or strkhmc="" Or strZrr="" Or strzywt="" Then
				Call JsAlert("请确认信息输入完整!请从正确的入口进入!","quality_list.asp")
			Else
				Call quality_Add()
			End If
		Case "change"
			If strhth="" Or strkhmc="" Or strZrr="" Or strzywt="" Or iid="" Then
				Call JsAlert("请确认信息输入完整!","")
			Else
				Call quality_Change()
			End If
		Case "delete"
			If iid=0 Then
				Call JsAlert("请确认从系统入口进入!","")
			Else
				strSql="delete from [quality] where id=" & iid
				Call xjweb.Exec(strSql, 0)
				Call JsAlert("外部质量信息删除成功","quality_list.asp")
			end if
		Case Else
			Call JsAlert("action="&action&", 请联系管理员!","quality_list.asp")
	End Select
	
	'验证外部信息完成情况
	Function quality_zt()
		If stryjcs <> "" Then strwczk = "分析中"		 
		If stryyfx <> "" Then strwczk = "制定措施中" 
		If strjzcs <> "" Then strwczk = "跟踪检查中" 
		If strlsqk <> "" Then strwczk = "验证中"  
		If stryzjl <> "" Then strwczk = "已闭环"  
	End Function
	
	'外部质量信息 入库
	Function quality_Add()
		strSql="select * from [quality]"
		Call xjweb.Exec("",-1)
		Call quality_zt()
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("khmc")=strkhmc
			Rs("lxr")=strlxr
			Rs("lxdh")=strlxdh
			Rs("hth")=strhth
			If strgzlh<>"" Then Rs("gzlh")=strgzlh
			Rs("jssj")=strjssj
			Rs("zrr")=strZrr
			Rs("zywt")=strzywt
			Rs("yjcs")=stryjcs
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl
			Rs("wczk")=strwczk
		Rs.Update
		Rs.Close
		Call JsAlert("外部质量信息添加成功","quality_add.asp")
	End Function

	'更改外部质量信息 入库	
	Function quality_Change()
		'检测ID号是否存在
		Set Rs=xjweb.Exec("select * from [quality] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("ID号 " & iid & " 问题分析不存在！","quality_list.asp")
			Rs.Close
			Exit Function
		End If
		Rs.Close

		strSql="select * from [quality] where id=" & iid
		Call xjweb.Exec("",-1)
		Call quality_zt()
		'strmsg="数据库操作"
		Rs.open strSql,conn,1,3
			Rs("khmc")=strkhmc
			Rs("lxr")=strlxr
			Rs("lxdh")=strlxdh
			Rs("hth")=strhth
			Rs("gzlh")=strgzlh
			Rs("jssj")=strjssj
			Rs("zrr")=strZrr
			Rs("zywt")=strzywt
			Rs("yjcs")=stryjcs
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl	
			Rs("wczk")=strwczk			
		Rs.update
		Rs.close

		Call JsAlert("外部质量信息更改成功","quality_list.asp")
	End Function
%>