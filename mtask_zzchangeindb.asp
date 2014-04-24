<!--#include file="include/conn.asp"-->
<%
	Call ChkAble(5)
	'本文件只负组长更改任务书的入库
	dim strmtjgr, strdxjgr, strgjjgr, strmtsjr, strdxsjr, strgjsjr, strmtshr, strdxshr, strgjshr, strmtbomr, strdxbomr, strlsh,strmtjgshr, strmtsjshr, strdxjgshr, strdxsjshr, strgjjgshr, strgjsjshr, strmtjgsj, strmtsjsj, strmtshsj, strmtbomsj, strmtjgshsj, strmtsjshsj, strdxjgsj, strdxsjsj, strdxshsj, strdxbomsj, strgjshsj, strgjjgshsj, strgjsjshsj, strdxjgshsj, strdxsjshsj, strgjjgsj, strgjsjsj, strgjfcr, strgjfcsj, strmtgysjr, strmtgysjjs, strdxgysjr, strdxgysjjs, strmtgyshr, strmtgyshjs, strdxgyshr, strdxgyshjs, strgjgysjr, strgjgysjks, strgjgysjjs, strgjgyshr, strgjgyshks, strgjgyshjs
	strmtjgr="" : strdxjgr="" : strgjjgr="" : strmtsjr="" : strdxsjr="" : strgjsjr=""
	strmtshr="" : strdxshr="" : strgjshr="" : strmtbomr="" : strdxbomr="" : strgjfcr="" : strgjfcsj=""
	strmtjgshr="" : strmtsjshr="" : strdxjgshr="" : strdxsjshr="" : strgjjgshr="" : strgjsjshr=""
	strmtgysjr="" : strdxgysjr="" : strmtgyshr="" : strdxgyshr="" :	strgjgysjr="" : strgjgyshr=""


	strmtjgsj=Replace(request("mtjgsj"),"."," ")
	strmtsjsj=Replace(request("mtsjsj"),"."," ")
	strmtshsj=Replace(request("mtshsj"),"."," ")
	strmtjgshsj=Replace(request("mtjgshsj"),"."," ")
	strmtsjshsj=Replace(request("mtsjshsj"),"."," ")
	strdxjgsj=Replace(request("dxjgsj"),"."," ")
	strdxsjsj=Replace(request("dxsjsj"),"."," ")
	strdxshsj=Replace(request("dxshsj"),"."," ")
	strmtbomsj=Replace(request("mtbomsj"),"."," ")
	strdxbomsj=Replace(request("mtbomsj"),"."," ")
	strdxjgshsj=Replace(request("dxjgshsj"),"."," ")
	strdxsjshsj=Replace(request("dxsjshsj"),"."," ")
	strgjjgsj=Replace(request("gjjgsj"),"."," ")
	strgjsjsj=Replace(request("gjsjsj"),"."," ")
	strgjshsj=Replace(request("gjshsj"),"."," ")
	strgjjgshsj=Replace(request("gjjgshsj"),"."," ")
	strgjsjshsj=Replace(request("gjsjshsj"),"."," ")
	strgjfcsj=Replace(request("gjfcsj"),"."," ")

	strmtgysjjs=Replace(request("mtgysjsj"),"."," ")
	strdxgysjjs=Replace(request("dxgysjsj"),"."," ")
	strgjgysjjs=Replace(request("gjgysjsj"),"."," ")
	strmtgyshjs=Replace(request("mtgyshsj"),"."," ")
	strdxgyshjs=Replace(request("dxgyshsj"),"."," ")
	strgjgyshjs=Replace(request("gjgyshsj"),"."," ")

	strmtjgr=request("mtjgr")
	strmtsjr=request("mtsjr")
	strmtshr=request("mtshr")
	strmtbomr=request("mtbomr")
	strdxjgr=request("dxjgr")
	strdxsjr=request("dxsjr")
	strdxshr=request("dxshr")
	strdxbomr=request("dxbomr")
	strgjjgr=request("gjjgr")
	strgjsjr=request("gjsjr")
	strgjshr=request("gjshr")
	strlsh=request("lsh")
	strmtjgshr=request("mtjgshr")
	strmtsjshr=request("mtsjshr")
	strdxjgshr=request("dxjgshr")
	strdxsjshr=request("dxsjshr")
	strgjjgshr=request("gjjgshr")
	strgjsjshr=request("gjsjshr")
	strgjfcr=request("gjfcr")

	strmtgysjr=request("mtgysjr")
	strdxgysjr=request("dxgysjr")
	strmtgyshr=request("mtgyshr")
	strdxgyshr=request("dxgyshr")
	strgjgysjr=request("gjgysjr")
	strgjgyshr=request("gjgyshr")

	If strlsh="" Then
		Call JsAlert("无法确定任务书的流水号,请从正规入口进入!", "")
	End If

	Set Rs=xjweb.Exec("select lsh from [mtask] where [lsh]='"&strlsh&"'",1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("流水号 【" & strlsh &  "】 任务书不存在!","")
	End If
	Rs.Close

'		Call JsAlert(strmtbomsj,"")

	strSql="select * from [mtask] where [lsh]='" & strlsh & "'"
	Call xjweb.Exec("",-1)
	strMsg=""
	Rs.open strsql,Conn,1,3
		if strmtjgr<>"" and strmtjgr<>rs("mtjgr") then rs("mtjgr")=strmtjgr : strmsg = strmsg & "更改模头结构人"
		if strmtjgsj<>"" and strmtjgr<>rs("mtjgjs") then rs("mtjgjs")=strmtjgsj : strmsg = strmsg & "更改模头结构结束时间"
		if strdxjgr<>"" and strdxjgr<>rs("dxjgr") then rs("dxjgr")=strdxjgr : strmsg = strmsg & "更改定型结构人"
		if strdxjgsj<>"" and strdxjgsj<>rs("dxjgjs") then rs("dxjgjs")=strdxjgsj : strmsg = strmsg & "更改定型结构结束时间"
		if strgjjgr<>"" and strgjjgr<>rs("gjjgr") then rs("gjjgr")=strgjjgr : strmsg = strmsg & "更改后共挤结构人"
		if strgjjgsj<>"" and strgjjgsj<>rs("gjjgjs") then rs("gjjgjs")=strgjjgsj : strmsg = strmsg & "更改后共挤结构结束时间"
		if strmtsjr<>"" and strmtsjr<>rs("mtsjr") then rs("mtsjr")=strmtsjr : strmsg = strmsg & "更改模头设计人"
		if strmtsjsj<>"" and strmtsjsj<>rs("mtsjjs") then rs("mtsjjs")=strmtsjsj : strmsg = strmsg & "更改模头设计结束时间"
		if strdxsjr<>"" and strdxsjr<>rs("dxsjr") then rs("dxsjr")=strdxsjr : strmsg = strmsg & "更改定型设计人"
		if strdxsjsj<>"" and strdxsjsj<>rs("dxsjjs") then rs("dxsjjs")=strdxsjsj : strmsg = strmsg & "更改定型设计结束时间"
		if strgjsjr<>"" and strgjsjr<>rs("gjsjr") then rs("gjsjr")=strgjsjr : strmsg = strmsg & "更改后共挤设计人"
		if strgjsjsj<>"" and strgjsjsj<>rs("gjsjjs") then rs("gjsjjs")=strgjsjsj : strmsg = strmsg & "更改后共挤设计结束时间"
		if strmtshr<>"" and strmtshr<>rs("mtshr") then rs("mtshr")=strmtshr : strmsg = strmsg & "更改模头审核人"
		if strmtshsj<>"" and strmtshsj<>rs("mtshjs") then rs("mtshjs")=strmtshsj : strmsg = strmsg & "更改模头审核结束时间"
		if strdxshr<>"" and strdxshr<>rs("dxshr") then rs("dxshr")=strdxshr : strmsg = strmsg & "更改定型审核人"
		if strdxshsj<>"" and strdxshsj<>rs("dxshjs") then rs("dxshjs")=strdxshsj : strmsg = strmsg & "更改定型审核结束时间"
		if strgjshr<>"" and strgjshr<>rs("gjshr") then rs("gjshr")=strgjshr : strmsg = strmsg & "更改后共挤审核人"
		if strgjshsj<>"" and strgjshsj<>rs("gjshjs") then rs("gjshjs")=strgjshsj : strmsg = strmsg & "更改后共挤审核结束时间"
		if strmtbomr<>"" and strmtbomr<>rs("mtbomr") then rs("mtbomr")=strmtbomr : strmsg = strmsg & "更改模头BOM人"
		if strmtbomsj<>"" and strmtbomsj<>rs("mtbomjs") then rs("mtbomjs")=strmtbomsj : strmsg = strmsg & "更改模头BOM结束时间"
		if strdxbomr<>"" and strdxbomr<>rs("dxbomr") then rs("dxbomr")=strdxbomr : strmsg = strmsg & "更改定型BOM人"
		if strdxbomsj<>"" and strdxbomsj<>rs("dxbomjs") then rs("dxbomjs")=strdxbomsj : strmsg = strmsg & "更改定型BOM结束时间"
		if strmtjgshr<>"" and strmtjgshr<>rs("mtjgshr") then rs("mtjgshr")=strmtjgshr : strmsg = strmsg & "更改模头结构确认人"
		if strmtjgshsj<>"" and strmtjgshsj<>rs("mtjgshjs") then rs("mtjgshjs")=strmtjgshsj : strmsg = strmsg & "更改模头结构确认结束时间"
		if strmtsjshr<>"" and strmtsjshr<>rs("mtsjshr") then rs("mtsjshr")=strmtsjshr : strmsg = strmsg & "更改模头设计确认人"
		if strmtsjshsj<>"" and strmtsjshsj<>rs("mtsjshjs") then rs("mtsjshjs")=strmtsjshsj : strmsg = strmsg & "更改模头设计确认结束时间"
		if strdxjgshr<>"" and strdxjgshr<>rs("dxjgshr") then rs("dxjgshr")=strdxjgshr : strmsg = strmsg & "更改定型结构确认人"
		if strdxjgshsj<>"" and strdxjgshsj<>rs("dxjgshjs") then rs("dxjgshjs")=strdxjgshsj : strmsg = strmsg & "更改定型结构确认结束时间"
		if strdxsjshr<>"" and strdxsjshr<>rs("dxsjshr") then rs("dxsjshr")=strdxsjshr : strmsg = strmsg & "更改定型设计确认人"
		if strdxsjshsj<>"" and strdxsjshsj<>rs("dxsjshjs") then rs("dxsjshjs")=strdxsjshsj : strmsg = strmsg & "更改定型设计确认结束时间"
		if strgjjgshr<>"" and strgjjgshr<>rs("gjjgshr") then rs("gjjgshr")=strgjjgshr : strmsg = strmsg & "更改后共挤结构确认人"
		if strgjjgshsj<>"" and strgjjgshsj<>rs("gjjgshjs") then rs("gjjgshjs")=strgjjgshsj : strmsg = strmsg & "更改后共挤结构确认结束时间"
		if strgjsjshr<>"" and strgjsjshr<>rs("gjsjshr") then rs("gjsjshr")=strgjsjshr : strmsg = strmsg & "更改后共挤设计确认人"
		if strgjsjshsj<>"" and strgjsjshsj<>rs("gjsjshjs") then rs("gjsjshjs")=strgjsjshsj : strmsg = strmsg & "更改后共挤设计确认结束时间"
		if strgjfcr<>"" and strgjfcr<>rs("gjshr") then rs("gjshr")=strgjfcr : strmsg = strmsg & "更改共挤复查人"
		if strgjfcsj<>"" and strgjfcsj<>rs("gjshjs") then rs("gjshjs")=strgjfcsj : strmsg = strmsg & "更改共挤复查结束时间"

		if strmtgysjr<>"" and strmtgysjr<>rs("mtgysjr") then rs("mtgysjr")=strmtgysjr : strmsg = strmsg & "更改模头工艺设计人"
		if strmtgysjjs<>"" and strmtgysjjs<>rs("mtgysjjs") then rs("mtgysjjs")=strmtgysjjs : strmsg = strmsg & "更改模头工艺设计结束时间"
		if strmtgyshr<>"" and strmtgyshr<>rs("mtgyshr") then rs("mtgyshr")=strmtgyshr : strmsg = strmsg & "更改模头工艺审核人"
		if strmtgyshjs<>"" and strmtgyshjs<>rs("mtgyshjs") then rs("mtgyshjs")=strmtgyshjs : strmsg = strmsg & "更改模头工艺审核结束时间"
		if strdxgysjr<>"" and strdxgysjr<>rs("dxgysjr") then rs("dxgysjr")=strdxgysjr : strmsg = strmsg & "更改定型工艺设计人"
		if strdxgysjjs<>"" and strdxgysjjs<>rs("dxgysjjs") then rs("dxgysjjs")=strdxgysjjs : strmsg = strmsg & "更改定型工艺设计结束时间"
		if strdxgyshr<>"" and strdxgyshr<>rs("dxgyshr") then rs("dxgyshr")=strdxgyshr : strmsg = strmsg & "更改定型工艺审核人"
		if strdxgyshjs<>"" and strdxgyshjs<>rs("dxgyshjs") then rs("dxgyshjs")=strdxgyshjs : strmsg = strmsg & "更改定型工艺审核结束时间"

		if strgjgysjr<>"" and strgjgysjr<>rs("gjgysjr") then rs("gjgysjr")=strgjgysjr : strmsg = strmsg & "更改共挤工艺设计人"
		if strgjgysjjs<>"" and strgjgysjjs<>rs("gjgysjjs") then rs("gjgysjjs")=strgjgysjjs : strmsg = strmsg & "更改共挤工艺设计结束时间"
		if strgjgyshr<>"" and strgjgyshr<>rs("gjgyshr") then rs("gjgyshr")=strgjgyshr : strmsg = strmsg & "更改共挤工艺审核人"
		if strgjgyshjs<>"" and strgjgyshjs<>rs("gjgyshjs") then rs("gjgyshjs")=strgjgyshjs : strmsg = strmsg & "更改共挤工艺审核结束时间"
	Rs.update
	Rs.Close

	If strMsg <> "" Then
		strMsg = "数据库操作: " & strMsg
		strSql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','更改任务书','"&strmsg&"','"&now()&"')"
		Call xjweb.Exec(strSql,0)
		Call JsAlert("流水号 【 " & strlsh & " 】任务书的责任人更改成功!", "mtask_zzchange.asp")
	Else
		Call JsAlert("您没有进行任何更改!","mtask_zzchange.asp")
	End If
%>