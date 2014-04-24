<!--#include file="include/conn.asp"-->
<%
	dim action, strrwlx, strrwlr, strzrr, izf, iid,strxldh,strxlxh,stryhdw,strgzyy,   strzbfa,strmjmc,strjssj,strzrfp,strylsh
	action="" : strrwlx="" : strrwlr="" : strzrr="" : izf=0 :  iid=0 :  strxldh="" : strxlxh="" : stryhdw="" : strgzyy="" :  strzbfa="" : strmjmc="":strjssj="":strzrfp="" : strylsh=""
	action=request("action")
	strrwlx=request("rwlx")
	strxldh=request("xldh")
	strxlxh=request("xlxh")
	stryhdw=request("yhdw")
	strzrfp=Request("zrfp")
	strylsh=trim(request("ylsh"))
	strgzyy=trim(request("gzyy"))
	strzbfa=trim(request("zbfa"))
	strmjmc=trim(request("mjmc"))
	iid=clng(request("id"))
	Call SelectT()

	'数据入库函数从这里开始
	select case action

		case "add"
			if strrwlx="" or strrwlr="" or strzrr="" or izf=0 then
				Call JsAlert("请确认信息输入完整!","")
			else
				call ftask_add()
			end if
		case "change"
			if strrwlx="" or strrwlr="" or strzrr="" or izf=0 or not(isnumeric(iid)) then
				Call JsAlert("请确认信息输入完整!","")
			else
				call ftask_change()
			end if
		case "delete"
			if not(isnumeric(iid)) then
				Call JsAlert("请确认从系统入口进入!","")
			else
				strSql="delete from [ftask] where id=" & iid
				call xjweb.Exec(strSql, 0)
				Call JsAlert("零星任务删除成功","ftask_list.asp")
			end if
		case else
			Call JsAlert("action="&action&", 请联系管理员!","ftask_list.asp")
	end select

	'零星任务书入库
	Function ftask_Add()
		Dim tmpSql, tmpRs, iGroup
		tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strzrr&"'"
		Set tmpRs=xjweb.Exec(tmpSql,1)
		If Not(tmpRs.Eof Or tmpRs.Bof) Then
			iGroup=tmpRs("user_group")
		Else
			iGroup=0
		End If
		tmpRs.Close
	
		strSql="select * from [ftask]"
		call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		rs.addnew
			rs("rwlx")=strrwlx
			if strxldh<>"" then rs("xldh")=strxldh end if
			rs("rwlr")=strrwlr
			rs("zrr")=strzrr
			rs("xz")=iGroup
			rs("zf")=izf
			rs("jssj")=strjssj
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update
		rs.close
		Call JsAlert("零星任务添加成功","ftask_add.asp")
	end function

	'更改零星任务入库
	function ftask_change()
		'检测流水号是否已存在
		set rs=xjweb.Exec("select * from [ftask] where id="&iid,1)
		if rs.eof or rs.bof then
			Call JsAlert("ID号 " & iid & " 零星任务不存在！","ftask_list.asp")
			Exit Function
		end if
		rs.close
		Dim tmpSql, tmpRs, iGroup
		tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strzrr&"'"
		Set tmpRs=xjweb.Exec(tmpSql,1)
		If Not(tmpRs.Eof Or tmpRs.Bof) Then
			iGroup=tmpRs("user_group")
		Else
			iGroup=0
		End If
		tmpRs.Close

		strSql="select * from [ftask] where id=" & iid
		call xjweb.Exec("",-1)
		'strmsg="数据库操作"
		rs.open strSql,conn,1,3
			rs("rwlx")=strrwlx
			if strxldh<>"" then rs("xldh")=strxldh end if
			rs("rwlr")=strrwlr
			rs("zrr")=strzrr
			rs("xz")=iGroup
			rs("zf")=izf
			rs("jssj")=strjssj
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update
		rs.close

		'sql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','更改任务书','"&strmsg&"','"&now()&"')"
		'call xjweb.Exec(sql,0)
		Call JsPrompt("零星任务更改成功")
	End Function
	function SelectT()
	'根据任务不同初始化数据
	if strrwlx="零星修理" then
	strrwlr="用户单位:"&stryhdw&"||模具名称:"&strmjmc&"||修理小号:"&strxlxh&"||故障现象与分析原因:"&strgzyy&"||准备采取方案:"&strzbfa&"||责任分配:"&strzrfp&"||原流水号:"&strylsh
	izf=Request("zf1")
	strzrr=Request("zrr1")
	strjssj=Request("psy1") & "年" & request("psm1") & "月" & request("psd1") & "日"
	else
	strrwlr=trim(request("rwlr"))
	izf=request("zf")
	strzrr=request("zrr")
	strjssj=request("psy") & "年" & request("psm") & "月" & request("psd") & "日"
	end if
	end function

%>