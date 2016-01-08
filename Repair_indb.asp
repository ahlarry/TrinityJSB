<!--#include file="include/conn.asp"-->
<%
	'2016-01-06 16:49
	'本文件只负责添加和更改修理任务书的入库a
	Dim action
	action=Request("action")
	'任务书所有变量初始化	
	Dim strlsh, strddh, strmh, strdwmc, strdmmc, strtslb, sngmjzf, sngmtbl, sngmtjgbl, sngdxjgbl, strbz, strjgzz, strsjzz, dtjhkssj, dtjgjssj, dtjhjssj
	Dim strmjxx, strrwlr, dtrwxdsj, dtlzrq, strlzr, ibomzf, itsdzf, itszf, itsxxzlzf
	strlsh=Trim(UCase(Request("lsh"))) : strddh=Trim(Request("ddh")) : strdwmc=Trim(Request("dwmc")) : strdmmc=Trim(Request("dmmc"))
	strmh=Trim(Request("mh")) : strtslb=Request("tslb") : sngmjzf=Request("mjzf") : sngmtbl=Request("mtbl") : sngmtjgbl=Request("mtjgbl")
	sngdxjgbl=Request("dxjgbl") : strbz=Request("bz") : strjgzz=Request("jgzz") : strsjzz=Request("sjzz")
	dtjhkssj=Request("jhkssj") : dtjhjssj=Request("jhjssj") : dtjgjssj=Request("jgjssj") :  strmjxx=Request("mjxx") : strrwlr=Request("rwlr")
	dtrwxdsj=now() : dtlzrq=now() : strlzr=session("userName")
	ibomzf=Request("bomzf") : itsdzf=Request("tsdzf") : itszf=Request("tszf") : itsxxzlzf=Request("tsxxzlzf")

	sngmjzf=NullToNum(sngmjzf)
	sngmtbl=NulltoNum(sngmtbl)
	sngmtjgbl=NulltoNum(sngmtjgbl)
	sngdxjgbl=NulltoNum(sngdxjgbl)
	ibomzf=NulltoNum(ibomzf)
	itsdzf=NulltoNum(itsdzf)
	itszf=NulltoNum(itszf)
	itsxxzlzf=NulltoNum(itsxxzlzf)

	'对待入库数据进行处理
	strMsg=""
	If strlsh="" Then strMsg="修理小号为空!<br>"
	If strddh="" Then strMsg=strMsg & "修理单号为空!<br>"
	If strdwmc="" Then strMsg=strMsg & "客户名称为空!<br>"
	If strmh=""  Then strMsg=strMsg & "原流水号为空!<br>"
	If strtslb=""  Then strMsg=strMsg & "调试类别不能为空!<br>"
	If sngmjzf=0 Then strMsg=strMsg & "模具总分为零!<br>"
	If sngmtjgbl=0 Then strMsg=strMsg & "模头结构比例为零!<br>"
	If sngdxjgbl=0 Then strMsg=strMsg & "定型结构比例为零!<br>"
	If strjgzz="" or strsjzz="" Then strMsg=strMsg & "组长没有选择!<br>"

	If strMsg <> "" and action <> "BzChan" Then
		infoTitle="数据不完整"
		infoContents=strMsg & "<br>点击<a href=""#"" onclick='history.go(-1);'>返回前页</a>重新输入"
		GotoPrompt()
	End If

	Call mtask_add()

	'添加任务书入库
	Function mtask_add()
		'检测流水号是否已存在
		Dim TmpRs
		Set TmpRs=xjweb.exec("select * from [mtask] where [lsh]='"&strlsh&"'",1)
		If Not(TmpRs.eof Or TmpRs.bof) Then
			If TmpRs("rwlr")<>"修理" Then 
				Call JsAlert("流水号 " & strlsh & " 为"& TmpRs("rwlr") &"任务，请更改流水号!","")
			else if isnull(TmpRs("mttsjs")) and isnull(TmpRs("dxtsjs")) Then
					strSql="select * from [mtask] where [lsh]='"& strlsh &"'"
					Call xjweb.exec("",-1)
					Rs.open strSql,Conn,1,3
					Rs("ddh")=strddh
					Rs("lsh")=strlsh
					Rs("dwmc")=strdwmc
					Rs("dmmc")=strdmmc
					Rs("mh")=strmh
					Rs("mjxx")=strmjxx
					Rs("rwlr")=strrwlr			
					If strtslb<>"" Then Rs("tslb")=strtslb
					If strbz<>"" Then Rs("bz")=strbz
					Rs("rwxdsj")=dtrwxdsj
					Rs("jhkssj")=dtjhkssj
					Rs("jhjgsj")=dtjgjssj
					Rs("jhjssj")=dtjhjssj
					Rs("jgzz")=strjgzz
					Rs("sjzz")=strsjzz
					Rs("lzr")=strlzr
					Rs("lzrq")=dtlzrq
					Rs("mjzf")=sngmjzf
					Rs("mtbl")=sngmtbl
					Rs("mtjgbl")=sngmtjgbl
					Rs("dxjgbl")=sngdxjgbl
					Rs("bomzf")=ibomzf
					Rs("tsdzf")=itsdzf
					Rs("tszf")=itszf
					Rs("tsxxzlzf")=itsxxzlzf
					Rs.update
					Rs.Close
					Call JsAlert("流水号 " & strlsh &  " 任务书更改成功!", "Repair_add.asp")	
				else
					Call JsAlert("流水号 " & strlsh & " 调试任务已完成，无法更改任务书!","")
				End If			
			End If
		else
			strSql="select * from [mtask]"
			Call xjweb.exec("",-1)
			Rs.open strSql,Conn,1,3
			Rs.AddNew
				Rs("ddh")=strddh
				Rs("lsh")=strlsh
				Rs("dwmc")=strdwmc
				Rs("dmmc")=strdmmc
				Rs("mh")=strmh
				Rs("mjxx")=strmjxx
				Rs("rwlr")=strrwlr			
				If strtslb<>"" Then Rs("tslb")=strtslb
				If strbz<>"" Then Rs("bz")=strbz
				Rs("rwxdsj")=dtrwxdsj
				Rs("jhkssj")=dtjhkssj
				Rs("jhjgsj")=dtjgjssj
				Rs("jhjssj")=dtjhjssj
				Rs("jgzz")=strjgzz
				Rs("sjzz")=strsjzz
				Rs("lzr")=strlzr
				Rs("lzrq")=dtlzrq
				Rs("mjzf")=sngmjzf
				Rs("mtbl")=sngmtbl
				Rs("mtjgbl")=sngmtjgbl
				Rs("dxjgbl")=sngdxjgbl
				Rs("bomzf")=ibomzf
				Rs("tsdzf")=itsdzf
				Rs("tszf")=itszf
				Rs("tsxxzlzf")=itsxxzlzf
			Rs.update
			Rs.Close
			Call JsAlert("任务书添加成功!", "Repair_add.asp")
		End If
		TmpRs.Close		
	End Function
%>
