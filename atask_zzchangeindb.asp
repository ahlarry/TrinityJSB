<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble(4)
	'本文件只负组长调试任务更改入库,如果完成同时更改分值库
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
		Call JsAlert("流水号不确定,无法确定具体的调试任务!\n请从正规入口进入!","atask_zzchange.asp")
	end if

	set rs=xjweb.Exec("select lsh from [mtask] where lsh='"&strlsh&"'",1)
	if rs.eof or rs.bof then
		Call JaAlert("流水号 【"&strlsh&"】 任务不存在!请核实!","atask_zzchange.asp")
	end if
	Rs.close

	strSql="select * from [mtask] where lsh='" & strlsh & "'"
	call xjweb.Exec("",-1)
	strmsg=""
	rs.open strSql,conn,1,3
		if strmttsdr<>"" and strmttsdr<>rs("mttsdr") then
			rs("mttsdr")=strmttsdr
			strmsg = strmsg & "更改模头调试单人"
			strSql="update [mantime] set zrr='"&strmttsdr&"' where lsh='"&strlsh&"' and rwlr='模头调试单'"
			call xjweb.Exec(strSql, 0)
		end if
		if strdxtsdr<>"" and strdxtsdr<>rs("dxtsdr") then 
			rs("dxtsdr")=strdxtsdr
			strmsg = strmsg & "更改定型调试单人"
			strSql="update [mantime] set zrr='"&strdxtsdr&"' where lsh='"&strlsh&"' and rwlr='定型调试单'"
			call xjweb.Exec(strSql, 0)
		end if
		if strmttsr<>"" and strmttsr<>rs("mttsr") then 
			rs("mttsr")=strmttsr
			strmsg = strmsg & "更改模头调试人"
			strSql="update [mantime] set zrr='"&strmttsr&"' where lsh='"&strlsh&"' and rwlr='模头调试'"
			call xjweb.Exec(strSql, 0)
		end if
		if strdxtsr<>"" and strdxtsr<>rs("dxtsr") then
			rs("dxtsr")=strdxtsr 
			strmsg = strmsg & "更改定型调试人"
			strSql="update [mantime] set zrr='"&strdxtsr&"' where lsh='"&strlsh&"' and rwlr='定型调试'"
			call xjweb.Exec(strSql, 0)
		end if
		if strmttsxxzlr<>"" and strmttsxxzlr<>rs("mttsxxzlr") then 
			rs("mttsxxzlr")=strmttsxxzlr
			strmsg = strmsg & "更改模头调试信息整理人"
			strSql="update [mantime] set zrr='"&strmttsxxzlr&"' where lsh='"&strlsh&"' and rwlr='模头调试信息整理'"
			call xjweb.Exec(strSql, 0)
		end if
		if strdxtsxxzlr<>"" and strdxtsxxzlr<>rs("dxtsxxzlr") then
			rs("dxtsxxzlr")=strdxtsxxzlr
			strmsg = strmsg & "更改定型调试信息整理人"
			strSql="update [mantime] set zrr='"&strdxtsxxzlr&"' where lsh='"&strlsh&"' and rwlr='定型调试信息整理'"
			call xjweb.Exec(strSql, 0)
		end if
	rs.update
	rs.close
	
	If strmsg<>"" Then
		strmsg="数据库操作:" & strmsg
		strSql="insert into [ims_log] (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','更改任务书','"&strmsg&"','"&now()&"')"
		call xjweb.Exec(strSql,0)
		Call JsAlert("流水号 【" & strlsh & "】 调试任务责任人更改成功!","atask_zzchange.asp")
	Else
		Call JsAlert("您没有对流水号 【" & strlsh & "】 调试任务进行任何更改!","atask_zzchange.asp")
	End If
%>