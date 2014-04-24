<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble("1,6")
	dim action, strtsyy, strtslr, btsps, strlsh, iid
	action="" : strtsyy="" : strtslr="" : btsps=false : strlsh=""
	action=request("action")
	strtsyy=request("tsyy")
	strtslr=trim(request("tslr"))
	btsps=request("tsps")
	strlsh=request("lsh")
	iid=request("id")

	'数据入库函数从这里开始
	select case action
		case "add"
			if strtsyy="" or strtslr="" or strlsh="" then
				Call JsAlert("请确认从正确入口进入并保证信息输入完整!","mtest_add.asp")
			else
				if btsps then
					Call mtestps_add()
				else
					Call mtest_add()
				end if
			end if
		case "change"
			if strtsyy="" or strtslr="" or not isnumeric(iid) then
				Call JsAlert("请确认从正确入口进入并保证信息输入完整!","mtest_add.asp")
			else
				Call mtest_change()
			end if
		case "delete"
			if not(isnumeric(iid)) then
				Call JsAlert("请确认从正确入口进入并保证信息输入完整!","mtest_add.asp")
			else
				Dim bPs
				strSql="delete from [ts_tsxx] where id=" & iid
				Set Rs=xjweb.Exec("Select lsh,ps from [ts_tsxx] where id=" & iid,1)
				'记录调试信息的相关信息
				strlsh=Rs("lsh")
				bPs=Rs("ps")
				Rs.Close
				Call xjweb.Exec(strSql, 0)
				'更改调试的次数
				strSql="update [ts_mould] set tscs="&xjweb.RsCount("[ts_tsxx] where lsh='"&strlsh&"' and not ps")&" where lsh='"&strlsh&"'"
				Call xjweb.Exec(strSql,0)
				If bPs Then
					Call JsAlert("调试评审信息删除成功!","mtest_display.asp?s_lsh="&strlsh&"")
				Else
					Call JsAlert("模具调试信息删除成功!","mtest_display.asp?s_lsh="&strlsh&"")
				End If
				response.end
			end if
		case else
			response.write "action=" & action
	end select

	'调试信息入库
	function mtest_add()
		strSql="select * from ts_tsxx"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		rs.addnew
			rs("lsh")=strlsh
			rs("tsyy")=strtsyy
			rs("tslr")=strtslr
			rs("tssj")=now()
			rs("tsr")=Session("userName")
		rs.update
		rs.close

		'将信息入ts_mould 表
		strSql="select * from [ts_mould] where lsh='"&strlsh&"'"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
			if isnull(rs("tskssj")) then rs("tskssj")=now()
			rs("tscs")=rs("tscs") + 1
			rs("tsgxsj")=now()
		rs.update
		rs.close
		Call JsAlert("流水号 【" & strlsh & "】 模具调试信息添加成功!","mtest_display.asp?s_lsh="&strlsh&"")
	end function

	function mtestps_add()
		strSql="select * from ts_tsxx"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		rs.addnew
			rs("lsh")=strlsh
			rs("tsyy")=strtsyy
			rs("tslr")=strtslr
			rs("tssj")=now()
			rs("tsr")=Session("userName")
			rs("ps")=true
		rs.update
		rs.close
		Call JsAlert("流水号 【" & strlsh & "】 模具调试评审添加成功!","mtest_display.asp?s_lsh="&strlsh&"")
	end function

	'更改零星任务入库
	function mtest_change()
		'检测流水号是否已存在
		set rs=xjweb.Exec("select * from [ts_tsxx] where id="&iid,1)
		if rs.eof or rs.bof then
			Call JsAlert("ID号为 〖"&iid&"〗 的调试信息不存在!","mtest_list.asp")
			exit function
		End if
		rs.close

		strSql="select * from ts_tsxx where id=" & iid
		Call xjweb.Exec("",-1)
		strmsg="数据库操作"
		rs.open strSql,conn,1,3
			rs("tsyy")=strtsyy
			rs("tslr")=strtslr
		rs.update
		rs.close

		'strSql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','更改任务书','"&strmsg&"','"&now()&"')"
		'Call xjweb.Exec(strSql,0)
		Call JsAlert("调试信息更改成功!","mtest_display.asp?s_lsh="&strlsh&"")
	end function
%>