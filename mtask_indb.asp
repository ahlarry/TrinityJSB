<!--#include file="include/conn.asp"-->
<%
'14:22 2007-1-6-星期六
	'本文件只负责添加和更改任务书的入库
	Dim action
	action=Request("action")

	Dim strlsh, strddh, strdwmc, strdmmc, strmh, strsbcj, strjcjxh, strmjxx, strrwlr, ifgbl, ifcbl
	Dim strmtjg, strdxjg, strsxjg, strsjtsl, strqjtsl, strqysd, strmtljcc, strrdogg, ijcfx, iqs
	Dim bpjrb, strjrbxs, strjrbcl, strjrbxx, strmjcl, strbz, dtjhjssj, strgjljcc
	Dim dtrwxdsj, strzz, strjsdb, strbm, strlzr, dtlzrq, bgjfs,bqhgj, strgjxs
	Dim sngmjzf, sngmtbl, ibomzf, itsdzf, itszf, itsxxzlzf, igjzf,sngmtjgbl, sngdxjgbl
	Dim strxcbh, strdxqg, strtslb, strcnts, bbtiao,strckdm,strfzxs, strmtrw, strdxrw
	Dim igjfs1, igjfs2, igjfs3, igjfs4, strjgzz, strsjzz, dtjhkssj, dtjgjssj, strdedm, strdefz, strdemt, strdedx

	'任务书所有变量初始化
	strlsh=Trim(Request("lsh")) : strddh=Trim(Request("ddh")) : strdwmc=Trim(Request("dwmc")) : strdmmc=Trim(Request("dmmc"))
	strmh=Trim(Request("mh")) : strsbcj=Trim(Request("sbcj")) : strjcjxh=Trim(Request("jcjxh")) : strmjxx=Request("mjxx")
	strrwlr=Request("rwlr") : strmtjg=Trim(Request("mtjg")) : strdxjg=Trim(Request("dxjg")) : strsxjg=Trim(Request("sxjg"))
	strsjtsl=Trim(Request("sjtsl")) : strqjtsl=Trim(Request("qjtsl")) : strqysd=Trim(Request("qysd")) : strmtljcc=Trim(Request("mtljcc"))
	strgjljcc=Trim(Request("gjljcc")) : strrdogg=Trim(Request("rdogg")) : ijcfx=Trim(Request("jcfx")) : iqs=Request("qs") : bpjrb=Request("pjrb")
	strjrbxs=Request("jrbxs") : strjrbcl=Request("jrbcl") : strjrbxx=Trim(Request("jrbxx")) : strmjcl=Trim(Request("mjcl"))
	strbz=Request("bz") : strckdm=Request("ckdm") : strfzxs=Request("fzxs") : dtjhkssj=Request("jhkssj") : dtjhjssj=Request("jhjssj")
	dtjgjssj=Request("jgjssj") :  dtrwxdsj=now() : dtlzrq=now() : sngmtjgbl=Request("mtjgbl") : sngdxjgbl=Request("dxjgbl")
	strzz=Request("zz") : strjsdb=Request("jsdb") : strbm=session("user_depart") : strlzr=session("userName")
	sngmjzf=Request("mjzf") : sngmtbl=Request("mtbl") : ibomzf=Request("bomzf") : itsdzf=Request("tsdzf")
	itszf=Request("tszf") : itsxxzlzf=Request("tsxxzlzf") : igjzf=Request("gjzf")
	strtslb=Request("tslb") : strxcbh=Request("xcbh") : strcnts=Request("cnts") : bbtiao=Request("beit")
	strjgzz=Request("jgzz") : strsjzz=Request("sjzz") : strdxqg=Request("dxqg")
	igjfs1=Request("ssgjf") : igjfs2=Request("qbfgjf")
	igjfs3=Request("qgjf") : igjfs4=Request("hgjf")
	strmtrw=Request("mtrw") : strdxrw=Request("dxrw") : strdedm=Request("dedm") : strdefz=Request("defz")		

	sngmjzf=NullToNum(sngmjzf)
	sngmtbl=NulltoNum(sngmtbl)
	sngmtjgbl=NulltoNum(sngmtjgbl)
	sngdxjgbl=NulltoNum(sngdxjgbl)
	ibomzf=NulltoNum(ibomzf)
	itsdzf=NulltoNum(itsdzf)
	itszf=NulltoNum(itszf)
	itsxxzlzf=NulltoNum(itsxxzlzf)
	igjzf=NulltoNum(igjzf)
	strdemt=NulltoNum(strdemt)
	strdedx=NulltoNum(strdedx)

	'初始化模具定额信息
	strSql="select * from [c_fzbl]"
	set rs=xjweb.Exec(strSql, 1)
	ifgbl=CSng(rs("fgbl"))
	ifcbl=CSng(rs("fcbl"))
	rs.close

	Select Case strmjxx
		Case "模头"
			strdefz=strdefz*0.4
			sngmtbl=100
		Case "定型"
			strdefz=strdefz*0.6
			sngmtbl=0
	End select	
	
	If igjfs2<>0 Then
		strdemt=strdefz*strfzxs*sngmtbl/100 + 400
	else
		strdemt=strdefz*strfzxs*sngmtbl/100
	End If
	strdedx=strdefz*strfzxs*(100-sngmtbl)/100
	If igjfs3<>0 Then
		strdemt=strdemt + 200*sngmtbl/100
		strdedx=strdedx + 200*(100-sngmtbl)/100
	End If	
	Select Case strmtrw
		Case ""
			strdemt=0
		Case "设计"
			strdemt=Round(strdemt,1)
		Case "复改"
			strdemt=Round(strdemt*ifgbl,1)
		Case "复查"
			strdemt=Round(strdemt*ifcbl,1)
	End select
	Select Case strdxrw
		Case ""
			strdedx=0
		Case "设计"
			strdedx=Round(strdedx,1)
		Case "复改"
			strdedx=Round(strdedx*ifgbl,1)
		Case "复查"
			strdedx=Round(strdedx*ifcbl,1)
	End select	

	'对待入库数据进行处理
	strMsg=""
	If strlsh="" Then strMsg="流水号为空!<br>"
	If strddh="" Then strMsg=strMsg & "订单号为空!<br>"
	If strdwmc="" Then strMsg=strMsg & "客户名称为空!<br>"
	If strdmmc="" Then strMsg=strMsg & "断面名称为空!<br>"
	If strmh=""  Then strMsg=strMsg & "模号为空!<br>"
	If strsbcj=""  Then strMsg=strMsg & "设备厂家为空!<br>"
	If strjcjxh=""  Then strMsg=strMsg & "挤出机型号为空!<br>"
	If strqysd=""  Then strMsg=strMsg & "牵引速度为空!<br>"
	If strckdm=""  Then strMsg=strMsg & "参考断面不能为空!<br>"
	If bbtiao=""  Then strMsg=strMsg & "厂内不调试时是否北调不能为空!<br>"
	If strcnts="true" and strtslb=""  Then strMsg=strMsg & "厂内调试时调试类别不能为空!<br>"
	If strxcbh=""  Then strMsg=strMsg & "型材壁厚为空!<br>"
	If strmjxx<>"模头" And strdxjg="" Then strMsg=strMsg & "定型结构为空!<br>"
	If strmjxx<>"模头" And strsxjg="" Then strMsg=strMsg & "水箱结构为空!<br>"
	If strmtljcc=""  Then strMsg=strMsg & "模头连接尺寸为空!<br>"
	If strrdogg=""  Then strMsg=strMsg & "热电偶规格为空!<br>"
	If strmjcl=""  Then strMsg=strMsg & "模具材料为空!<br>"
	If sngmjzf=0 Then strMsg=strMsg & "模具总分为零!<br>"
	'If sngmtbl=0 Then strMsg=strMsg & "模头比例为零!<br>"
	If sngmtjgbl=0 Then strMsg=strMsg & "模头结构比例为零!<br>"
	If sngdxjgbl=0 Then strMsg=strMsg & "定型结构比例为零!<br>"
	If strdemt=0 and strdedx=0 Then strMsg=strMsg & "模头和定型定额不能同时为0!<br>"
	If strzz=""  Then
		If strjgzz="" or strsjzz="" Then strMsg=strMsg & "组长没有选择!<br>"
	End If
	If strjsdb=""  Then strMsg=strMsg & "技术代表没有选择!<br>"

	If strMsg <> "" and action <> "BzChan" Then
		infoTitle="数据不完整"
		infoContents=strMsg & "<br>点击<a href=""#"" onclick='history.go(-1);'>返回前页</a>重新输入"
		GotoPrompt()
	End If

	If Not ChkStr(strlsh) and action <> "BzChan" Then Call JsAlert("流水号中含有非法字符！\n如空格、回车、单引号、双引号等！","")

	'由组长得出是第几组
'	Dim igroup
'	igroup=0
'	strSql="select [user_group] from [ims_user] where [user_name]='"&strzz&"'"
'	Set Rs=xjweb.exec(strSql, 1)
'	igroup = Rs("user_group")

	'数据入库函数从这里开始
	Select Case action
		Case "add"
			Call mtask_add()
		Case "change"
			Call mtask_change()
		Case "BzChan"
			Call Bz_Chan()
		Case else
			response.write "action=" & action
	End select

	'添加任务书入库
	Function mtask_add()
		'检测流水号是否已存在
		Set Rs=xjweb.exec("select lsh from [mtask] where [lsh]='"&strlsh&"'",1)
		If Not(Rs.eof Or Rs.bof) Then
			Call JsAlert("流水号 " & strlsh & " 任务书已存在!请更改流水号!","")
			Exit Function
		End If
		Rs.Close
		'Response.End
		strSql="select * from [mtask]"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Rs.AddNew
			Rs("ddh")=strddh
			Rs("lsh")=strlsh
			Rs("dwmc")=strdwmc
			Rs("dmmc")=strdmmc
			Rs("mh")=strmh
			Rs("sbcj")=strsbcj
			Rs("jcjxh")=strjcjxh
			If strmtjg<>"" Then Rs("mtjg")=strmtjg
			If strdxjg<>"" Then Rs("dxjg")=strdxjg
			If strsxjg<>"" Then Rs("sxjg")=strsxjg
			Rs("sjtsl")=strsjtsl
			Rs("qjtsl")=strqjtsl
			Rs("qysd")=strqysd
			Rs("mtljcc")=strmtljcc
			Rs("gjljcc")=strgjljcc
			Rs("rdogg")=strrdogg
			Rs("mjxx")=strmjxx
			Rs("rwlr")=strrwlr			
			Rs("mtrw")=strmtrw
			Rs("dxrw")=strdxrw
			Rs("ckdm")=strckdm
			Rs("dedm")=strdedm
			Rs("demt")=strdemt
			Rs("dedx")=strdedx
			Rs("fzxs")=strfzxs
			Rs("cnts")=strcnts
			Rs("beit")=bbtiao
			If strtslb<>"" Then Rs("tslb")=strtslb
			Rs("xcbh")=strxcbh
			Rs("dxqg")=strdxqg
			Rs("jcfx")=ijcfx
			Rs("qs")=iqs
			Rs("pjrb")=bpjrb
			Rs("jrbxs")=strjrbxs
			Rs("jrbcl")=strjrbcl
			If strjrbxx<>"" Then Rs("jrbxx")=strjrbxx
			Rs("mjcl")=strmjcl
			If strbz<>"" Then Rs("bz")=strbz
			Rs("rwxdsj")=dtrwxdsj
			Rs("jhkssj")=dtjhkssj
			Rs("jhjgsj")=dtjgjssj
			Rs("jhjssj")=dtjhjssj
'			Rs("zz")=strzz
			Rs("jgzz")=strjgzz
			Rs("sjzz")=strsjzz
'			Rs("group")=igroup
			Rs("jsdb")=strjsdb
			Rs("bm")=strbm
			Rs("lzr")=strlzr
			Rs("lzrq")=dtlzrq
'			Rs("gjfs")=bgjfs
			Rs("SSGJ")=igjfs1
			Rs("QBFGJ")=igjfs2
			Rs("QGJ")=igjfs3
			Rs("HGJ")=igjfs4
'			Rs("qhgj")=bqhgj
			Rs("mjzf")=sngmjzf
			Rs("mtbl")=sngmtbl
			Rs("mtjgbl")=sngmtjgbl
			Rs("dxjgbl")=sngdxjgbl
			Rs("bomzf")=ibomzf
			Rs("tsdzf")=itsdzf
			Rs("tszf")=itszf
			Rs("tsxxzlzf")=itsxxzlzf
			Rs("gjzf")=igjzf
		Rs.update
		Rs.Close
		Call JsAlert("任务书添加成功!", "mtask_add.asp")
		Response.End
	End Function

	'更改任务书入库
	Function mtask_change()
		'检测流水号是否已存在
		Dim iid
		iid=Request("id")
		Set Rs=xjweb.Exec("select lsh from [mtask] where [lsh]='"&strlsh&"' And id<>"&iid&" ",1)
		If Not(Rs.eof Or Rs.bof) Then
			Call JsAlert("流水号 " & strlsh & " 任务书已存在! 请更改流水号!","")
			Exit Function
		End If
		Rs.Close

		strSql="select * from [mtask] where [id]=" & iid
		Call xjweb.exec("",-1)
		strMsg="更改任务书"
		Rs.open strSql,Conn,1,3
			Rs("ddh")=strddh
			Rs("lsh")=strlsh
			Rs("dwmc")=strdwmc
			Rs("dmmc")=strdmmc
			Rs("mh")=strmh
			Rs("sbcj")=strsbcj
			Rs("jcjxh")=strjcjxh
			If strmtjg<>"" Then Rs("mtjg")=strmtjg
			If strdxjg<>"" Then Rs("dxjg")=strdxjg
			If strsxjg<>"" Then Rs("sxjg")=strsxjg
			Rs("sjtsl")=strsjtsl
			Rs("qjtsl")=strqjtsl
			Rs("qysd")=strqysd
			Rs("mtljcc")=strmtljcc
			Rs("gjljcc")=strgjljcc
			Rs("rdogg")=strrdogg
			Rs("mjxx")=strmjxx
			Rs("rwlr")=strrwlr			
			Rs("mtrw")=strmtrw
			Rs("dxrw")=strdxrw
			Rs("ckdm")=strckdm
			Rs("dedm")=strdedm
			Rs("demt")=strdemt
			Rs("dedx")=strdedx
			Rs("fzxs")=strfzxs
			Rs("cnts")=strcnts
			Rs("beit")=bbtiao
			If strtslb<>"" Then Rs("tslb")=strtslb
			Rs("xcbh")=strxcbh
			Rs("dxqg")=strdxqg
			Rs("jcfx")=ijcfx
			Rs("qs")=iqs
			Rs("pjrb")=bpjrb
			Rs("jrbxs")=strjrbxs
			Rs("jrbcl")=strjrbcl
			If strjrbxx<>"" Then Rs("jrbxx")=strjrbxx
			Rs("mjcl")=strmjcl
			If strbz<>"" Then Rs("bz")=strbz
			Rs("rwxdsj")=dtrwxdsj
			Rs("jhkssj")=dtjhkssj
			Rs("jhjgsj")=dtjgjssj
			Rs("jhjssj")=dtjhjssj
			Rs("zz")=strzz
			Rs("jgzz")=strjgzz
			Rs("sjzz")=strsjzz
'			Rs("group")=igroup
			Rs("jsdb")=strjsdb
			Rs("bm")=strbm
			Rs("lzr")=strlzr
			Rs("lzrq")=dtlzrq
			Rs("gjfs")=0
			Rs("qhgj")=0
			Rs("SSGJ")=igjfs1
			Rs("QBFGJ")=igjfs2
			Rs("QGJ")=igjfs3
			Rs("HGJ")=igjfs4
			Rs("mjzf")=sngmjzf
			Rs("mtbl")=sngmtbl
			Rs("mtjgbl")=sngmtjgbl
			Rs("dxjgbl")=sngdxjgbl
			Rs("bomzf")=ibomzf
			Rs("tsdzf")=itsdzf
			Rs("tszf")=itszf
			Rs("tsxxzlzf")=itsxxzlzf
			Rs("gjzf")=igjzf
		Rs.update
		Rs.Close
		Call JsAlert("流水号 " & strlsh &  " 任务书更改成功!", "mtask_change.asp")
		Response.End
	End Function

	'更改备注
	Function Bz_Chan()
		'检测流水号是否已存在
		Dim iid
		iid=Request("id")
		strSql="select * from [mtask] where [id]=" & iid
		Call xjweb.exec("",-1)
		strMsg="更改任务书"
		Rs.open strSql,Conn,1,3
			If strbz<>"" Then Rs("bz")=strbz
		Rs.update
		Rs.Close
		Call JsAlert("备注更改成功!", "mtask_change.asp")
		Response.End
	End Function
%>
