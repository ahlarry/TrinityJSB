<!--#include file="include/conn.asp"-->
<%
'11:22 2007-5-30-星期三
	'本文件只负责分配任务书的入库
	Call ChkPageAble("3,4")
	dim strfplr, strzrr, strlsh, strSql2, Rs2, strhth, imtsjshf, idxsjshf
	strSql2="" : imtsjshf=0 : idxsjshf=0
	strfplr=request("fplr")
	strzrr=request("zrr")
	strlsh=request("lsh")
	strhth=request("hth")
	if strfplr="" or (strzrr="" and instr(strfplr,"开始")>0) or (strlsh="" and strhth="") then
		Call JsAlert("分配任务书信息不够! \n任务内容为空或没有责任人!","")
	end if
	dim bok
	bok=true
	if strlsh<>"" Then
		strSql2="select * from [mtask] where lsh='"&strlsh&"'"
		Set Rs2=xjweb.Exec(strSql2,1)
	else
		strSql2="select * from [jsdb] where hth='"&strhth&"'"
		Set Rs2=xjweb.Exec(strSql2,1)
	End if
	strSql=""
	select case strfplr
		case "开始模头结构"
			strSql="update [mtask] set mtjgr='"&strzrr&"', mtjgks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头结构", strzrr, now())

		case "开始定型结构"
			strSql="update [mtask] set dxjgr='"&strzrr&"', dxjgks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型结构", strzrr, now())

		case "开始后共挤结构"
			strSql="update [mtask] set gjjgr='"&strzrr&"', gjjgks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "后共挤结构", strzrr, now())

		case "开始全套结构"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtjgr='"&strzrr&"', mtjgks='"&now()&"', dxjgr='"&strzrr&"', gjjgr='"&strzrr&"', gjjgks='"&now()&"', dxjgks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头结构", strzrr, now())
				call taskstart(strlsh, "定型结构", strzrr, now())
				call taskstart(strlsh, "后共挤结构", strzrr, now())
			else
				strSql="update [mtask] set mtjgr='"&strzrr&"', mtjgks='"&now()&"', dxjgr='"&strzrr&"', dxjgks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头结构", strzrr, now())
				call taskstart(strlsh, "定型结构", strzrr, now())
			end if

		case "开始模头结构确认"
			strSql="update [mtask] set mtjgshr='"&strzrr&"', mtjgshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头结构确认", strzrr, now())

		case "开始定型结构确认"
			strSql="update [mtask] set dxjgshr='"&strzrr&"', dxjgshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型结构确认", strzrr, now())

		case "开始后共挤结构确认"
			strSql="update [mtask] set gjjgshr='"&strzrr&"', gjjgshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "后共挤结构确认", strzrr, now())

		case "开始全套结构确认"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtjgshr='"&strzrr&"', mtjgshks='"&now()&"', dxjgshr='"&strzrr&"', gjjgshr='"&strzrr&"', dxjgshks='"&now()&"', gjjgshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头结构确认", strzrr, now())
				call taskstart(strlsh, "定型结构确认", strzrr, now())
				call taskstart(strlsh, "后共挤结构确认", strzrr, now())
			else
				strSql="update [mtask] set mtjgshr='"&strzrr&"', mtjgshks='"&now()&"', dxjgshr='"&strzrr&"', dxjgshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头结构确认", strzrr, now())
				call taskstart(strlsh, "定型结构确认", strzrr, now())
			end if

		case "开始模头设计"
			strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头设计", strzrr, now())

		case "开始定型设计"
			strSql="update [mtask] set dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型设计", strzrr, now())

		case "开始后共挤设计"
			strSql="update [mtask] set gjsjr='"&strzrr&"', gjsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "后共挤设计", strzrr, now())

		case "开始全套设计"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', gjsjr='"&strzrr&"', gjsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头设计", strzrr, now())
				call taskstart(strlsh, "定型设计", strzrr, now())
				call taskstart(strlsh, "后共挤设计", strzrr, now())
			else
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头设计", strzrr, now())
				call taskstart(strlsh, "定型设计", strzrr, now())
			end if

		case "开始模头设计确认"
			strSql="update [mtask] set mtsjshr='"&strzrr&"', mtsjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头设计确认", strzrr, now())

		case "开始定型设计确认"
			strSql="update [mtask] set dxsjshr='"&strzrr&"', dxsjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型设计确认", strzrr, now())

		case "开始后共挤设计确认"
			strSql="update [mtask] set gjsjshr='"&strzrr&"', gjsjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "后共挤设计确认", strzrr, now())

		case "开始全套设计确认"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtsjshr='"&strzrr&"', mtsjshks='"&now()&"', dxsjshr='"&strzrr&"', gjsjshr='"&strzrr&"', dxsjshks='"&now()&"', gjsjshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头设计确认", strzrr, now())
				call taskstart(strlsh, "定型设计确认", strzrr, now())
				call taskstart(strlsh, "后共挤设计确认", strzrr, now())
			else
				strSql="update [mtask] set mtsjshr='"&strzrr&"', mtsjshks='"&now()&"', dxsjshr='"&strzrr&"', dxsjshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头设计确认", strzrr, now())
				call taskstart(strlsh, "定型设计确认", strzrr, now())
			end if

		case "开始模头审核"
			strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头审核", strzrr, now())

		case "开始定型审核"
			strSql="update [mtask] set dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型审核", strzrr, now())

		case "开始后共挤审核"
			strSql="update [mtask] set gjshr='"&strzrr&"', gjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "后共挤审核", strzrr, now())

		case "开始全套审核"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', gjshr='"&strzrr&"', gjshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头审核", strzrr, now())
				call taskstart(strlsh, "定型审核", strzrr, now())
				call taskstart(strlsh, "后共挤审核", strzrr, now())
			else
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头审核", strzrr, now())
				call taskstart(strlsh, "定型审核", strzrr, now())
			end if

'		case "全套工艺设计"
'			strSql="update [mtask] set mtgysjr='"&strzrr&"', mtgysjks='"&now()&"', mtgysjjs='"&now()&"', dxgysjr='"&strzrr&"', dxgysjks='"&now()&"', dxgysjjs='"&now()&"', gjgysjr='"&strzrr&"', gjgysjks='"&now()&"', gjgysjjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)
'		case "模头工艺设计"
'			strSql="update [mtask] set  mtgysjr='"&strzrr&"', mtgysjks='"&now()&"', mtgysjjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)
'		case "定型工艺设计"
'			strSql="update [mtask] set dxgysjr='"&strzrr&"', dxgysjks='"&now()&"', dxgysjjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)
'		case "共挤工艺设计"
'			strSql="update [mtask] set gjgysjr='"&strzrr&"', gjgysjks='"&now()&"', gjgysjjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)
'		case "全套工艺审核"
'			strSql="update [mtask] set mtgyshr='"&strzrr&"', mtgyshks='"&now()&"', mtgyshjs='"&now()&"', dxgyshr='"&strzrr&"', dxgyshks='"&now()&"', dxgyshjs='"&now()&"', gjgyshr='"&strzrr&"', gjgyshks='"&now()&"', gjgyshjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)
'		case "模头工艺审核"
'			strSql="update [mtask] set mtgyshr='"&strzrr&"', mtgyshks='"&now()&"', mtgyshjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)
'		case "定型工艺审核"
'			strSql="update [mtask] set dxgyshr='"&strzrr&"', dxgyshks='"&now()&"', dxgyshjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)
'		case "共挤工艺审核"
'			strSql="update [mtask] set gjgyshr='"&strzrr&"', gjgyshks='"&now()&"', gjgyshjs='"&now()&"' where lsh='"&strlsh&"'"
'			call xjweb.Exec(strSql, 0)

		case "开始模头BOM"
			strSql="update [mtask] set mtbomr='"&strzrr&"', mtbomks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头BOM", strzrr, now())

		case "开始定型BOM"
			strSql="update [mtask] set dxbomr='"&strzrr&"', dxbomks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型BOM", strzrr, now())

		case "开始全套BOM"
			strSql="update [mtask] set mtbomr='"&strzrr&"', mtbomks='"&now()&"', dxbomr='"&strzrr&"', dxbomks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头BOM", strzrr, now())
			call taskstart(strlsh, "定型BOM", strzrr, now())

		case "开始模头复改"
			strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头复改", strzrr, now())

		case "开始定型复改"
			strSql="update [mtask] set dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型复改", strzrr, now())

		case "开始后共挤复改"
			strSql="update [mtask] set gjsjr='"&strzrr&"', gjsjks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "后共挤复改", strzrr, now())

		case "开始全套复改"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', gjsjr='"&strzrr&"', gjsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头复改", strzrr, now())
				call taskstart(strlsh, "定型复改", strzrr, now())
				call taskstart(strlsh, "后共挤复改", strzrr, now())
			else
				strSql="update [mtask] set mtsjr='"&strzrr&"', mtsjks='"&now()&"', dxsjr='"&strzrr&"', dxsjks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头复改", strzrr, now())
				call taskstart(strlsh, "定型复改", strzrr, now())
			end if

		case "开始模头复查"
			strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "模头复查", strzrr, now())

		case "开始定型复查"
			strSql="update [mtask] set dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "定型复查", strzrr, now())

		case "开始后共挤复查"
			strSql="update [mtask] set gjshr='"&strzrr&"', gjshks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strlsh, "后共挤复查", strzrr, now())

		case "开始全套复查"
			if (Rs2("gjfs")=3) and (Rs2("qhgj")=2) Then
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', gjshr='"&strzrr&"', gjshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头复查", strzrr, now())
				call taskstart(strlsh, "定型复查", strzrr, now())
				call taskstart(strlsh, "后共挤复查", strzrr, now())
			else
				strSql="update [mtask] set mtshr='"&strzrr&"', mtshks='"&now()&"', dxshr='"&strzrr&"', dxshks='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskstart(strlsh, "模头复查", strzrr, now())
				call taskstart(strlsh, "定型复查", strzrr, now())
			end if

		case "结束模头结构"
			strSql="update [mtask] set mtjgjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头结构")

		case "结束定型结构"
			strSql="update [mtask] set dxjgjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型结构")

		case "结束后共挤结构"
			strSql="update [mtask] set gjjgjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "后共挤结构")

		case "结束全套结构"
			if not(isnull(Rs2("gjjgks"))) Then
				strSql="update [mtask] set mtjgjs='"&now()&"', dxjgjs='"&now()&"', gjjgjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头结构")
				call taskend(strlsh, "定型结构")
				call taskend(strlsh, "后共挤结构")
			else
				strSql="update [mtask] set mtjgjs='"&now()&"', dxjgjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头结构")
				call taskend(strlsh, "定型结构")
			End if

		case "结束模头结构确认"
			strSql="update [mtask] set mtjgshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头结构确认")

		case "结束定型结构确认"
			strSql="update [mtask] set dxjgshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型结构确认")

		case "结束后共挤结构确认"
			strSql="update [mtask] set gjjgshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "后共挤结构确认")

		case "结束全套结构确认"
			if not(isnull(Rs2("gjjgshks"))) Then
				strSql="update [mtask] set mtjgshjs='"&now()&"', dxjgshjs='"&now()&"', gjjgshjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头结构确认")
				call taskend(strlsh, "定型结构确认")
				call taskend(strlsh, "后共挤结构确认")
			else
				strSql="update [mtask] set mtjgshjs='"&now()&"', dxjgshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头结构确认")
				call taskend(strlsh, "定型结构确认")
			End if

		case "结束模头设计"
			strSql="update [mtask] set mtsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头设计")

		case "结束定型设计"
			strSql="update [mtask] set dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型设计")

		case "结束后共挤设计"
			strSql="update [mtask] set gjsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "后共挤设计")

		case "结束全套设计"
			if not(isnull(Rs2("gjsjks"))) Then
				strSql="update [mtask] set mtsjjs='"&now()&"', gjsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头设计")
				call taskend(strlsh, "定型设计")
				call taskend(strlsh, "后共挤设计")
			else
				strSql="update [mtask] set mtsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头设计")
				call taskend(strlsh, "定型设计")
			End if

		case "结束模头设计确认"
			strSql="update [mtask] set mtsjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头设计确认")

		case "结束定型设计确认"
			strSql="update [mtask] set dxsjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型设计确认")

		case "结束后共挤设计确认"
			strSql="update [mtask] set gjsjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "后共挤设计确认")

		case "结束全套设计确认"
			if not(isnull(Rs2("gjsjshks"))) Then
				strSql="update [mtask] set mtsjshjs='"&now()&"', dxsjshjs='"&now()&"', gjsjshjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头设计确认")
				call taskend(strlsh, "定型设计确认")
				call taskend(strlsh, "后共挤设计确认")
			else
				strSql="update [mtask] set mtsjshjs='"&now()&"', dxsjshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头设计确认")
				call taskend(strlsh, "定型设计确认")
			End if

		case "结束模头审核"
			strSql="update [mtask] set mtshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头审核")

		case "结束定型审核"
			strSql="update [mtask] set dxshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型审核")

		case "结束后共挤审核"
			strSql="update [mtask] set gjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "后共挤审核")

		case "结束全套审核"
			if not(isnull(Rs2("gjshks"))) Then
				strSql="update [mtask] set mtshjs='"&now()&"', gjshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头审核")
				call taskend(strlsh, "定型审核")
				call taskend(strlsh, "后共挤审核")
			else
				strSql="update [mtask] set mtshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头审核")
				call taskend(strlsh, "定型审核")
			End if

		case "结束模头BOM"
			strSql="update [mtask] set mtbomjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头BOM")

		case "结束定型BOM"
			strSql="update [mtask] set dxbomjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型BOM")

		case "结束全套BOM"
			strSql="update [mtask] set mtbomjs='"&now()&"', dxbomjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头BOM")
			call taskend(strlsh, "定型BOM")

		case "结束模头复改"
			strSql="update [mtask] set mtsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头复改")

		case "结束定型复改"
			strSql="update [mtask] set dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型复改")

		case "结束后共挤复改"
			strSql="update [mtask] set gjsjjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "后共挤复改")

		case "结束全套复改"
			if not(isnull(Rs2("gjsjks"))) Then
				strSql="update [mtask] set mtsjjs='"&now()&"', gjsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头复改")
				call taskend(strlsh, "定型复改")
				call taskend(strlsh, "后共挤复改")
			else
				strSql="update [mtask] set mtsjjs='"&now()&"', dxsjjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头复改")
				call taskend(strlsh, "定型复改")
			End if

		case "结束模头复查"
			strSql="update [mtask] set mtshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模头复查")

		case "结束定型复查"
			strSql="update [mtask] set dxshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "定型复查")

		case "结束后共挤复查"
			strSql="update [mtask] set gjshjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "后共挤复查")

		case "结束全套复查"
			if not(isnull(Rs2("gjsjks"))) Then
				strSql="update [mtask] set mtshjs='"&now()&"', gjshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头复查")
				call taskend(strlsh, "定型复查")
				call taskend(strlsh, "后共挤复查")
			else
				strSql="update [mtask] set mtshjs='"&now()&"', dxshjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call taskend(strlsh, "模头复查")
				call taskend(strlsh, "定型复查")
			End if

		case "结束复审"
			strSql="update [mtask] set psjl='"&request("psjl")&"', fsr='"&strzrr&"', fsjs='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strlsh, "模具复审")
			call sendmsg("伍新安", web_info(0), "模具全套结束", "<a href=mtask_display.asp?s_lsh="&strlsh&" target=""_blank"">流水号 <b>"&strlsh&"</b> 已结束复审，请求准予本模具全套结束。</a>")
			Call JsAlert("流水号 【" & strlsh & "】 任务书组长分配部分完成!","mtask_assign.asp")

		case "全套结束"
			if not bmtaskend("sjjssj", strlsh) then
				strSql="update [mtask] set sjjssj='"&request("psd")&"', psjl='"&request("zrpsjl")&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql, 0)
				call fentodb(strlsh)
				if datediff("d",Rs2("sjjssj"),Rs2("jhjssj"))<0 Then
'					Call JsAlert("计划时间:【" & Rs2("jhjssj") & "】 实际时间; 【"& Rs2("sjjssj") & "】","")
					if Rs2("jgzz")<>Rs2("sjzz") Then
						call KptoDb(strlsh,Rs2("jgzz"),Rs2("sjjssj"),"设计延迟")
						call KptoDb(strlsh,Rs2("sjzz"),Rs2("sjjssj"),"设计延迟")
					else
						call KptoDb(strlsh,Rs2("jgzz"),Rs2("sjjssj"),"设计延迟")
					End If
				End If
			end if
			Call JsAlert("流水号 【" & strlsh & "】 任务书全部完成!","mtask_assign.asp")

		case "开始技术代表设计"
			strSql="update [jsdb] set sjr='"&strzrr&"', sjkssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strhth, "技术代表设计", strzrr, now())
		case "结束技术代表设计"
			strSql="update [jsdb] set sjjssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strhth, "技术代表设计")

		case "开始技术代表审核"
			strSql="update [jsdb] set shr='"&strzrr&"', shkssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskstart(strhth, "技术代表审核", strzrr, now())
		case "结束技术代表审核"
			strSql="update [jsdb] set shjssj='"&now()&"' where hth='"&strhth&"'"
			call xjweb.Exec(strSql, 0)
			call taskend(strhth, "技术代表审核")
			call jsdbfz(strhth,Rs2("khmc"),Rs2("rwnr"),Rs2("jcf"),Rs2("jhjssj"),Rs2("sjr"),Rs2("sjjssj"),Rs2("shr"),Rs2("shjssj"))
		case else
			bok=false
			Call JsAlert("请记下以下内容并联系管理员:\n\n"& strfplr & "\n\n (系统异常)请联系管理员","mtask_assign.asp")
	end select
	if bok and strfplr<>"结束复审" and strfplr<>"全套结束" then
		Call JsAlert("任务书分配成功!","mtask_assign.asp?s_lsh="&strlsh&"&s_hth="&strhth&"")
	end if

function taskstart(lsh, rwlr, zrr, kssj)
	strSql="insert into [mtask_cur] (lsh, rwlr, zrr, kssj) values ('"&lsh&"', '"&rwlr&"', '"&zrr&"','"&kssj&"')"
	call xjweb.Exec(strSql, 0)
end function

function taskend(lsh, rwlr)
	strSql="delete from [mtask_cur] where lsh='"&lsh&"' and rwlr='"&rwlr&"'"
	call xjweb.Exec(strSql, 0)
end function

function bmtaskend(trs, lsh)		'防止分值再次入库
	bmtaskend=true
	strSql="select "&trs&" from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	if isnull(rs(trs)) then bmtaskend=false
	rs.close
end function

Function FenToDB(lsh)
	'将分值写入分值库
	Dim mjfz, mtfz, dxfz, gjfz, bomfz, ijgbl, isjbl, ishbl, ifgbl, ifgshbl, ifcbl, ijc, ijc2, imtjgbl, idxjgbl, ijgshbl, iljshbl, iwcsj, mtgjf, dxgjf, ssgjf, qbfgjf, qgjf, hgjf
	Dim igysjxs, igysjsh, igyfcxs, igyfcsh, igyfgxs, igyfgsh, iGroup, tmpSql, tmpRs, itsdfz, sngmtbl, ifsxs
	mtgjf=0 : dxgjf=0 : ifsxs=0
	'ijc===奖惩分值
	strSql="select * from [c_fzbl]"
	set rs=xjweb.Exec(strSql, 1)
	imtjgbl=CSng(rs("mtjgbl"))
	idxjgbl=CSng(rs("dxjgbl"))
	ijgshbl=CSng(rs("jgshbl"))
	iljshbl=CSng(rs("ljshbl"))
	ishbl=CSng(rs("shbl"))
	ifgbl=CSng(rs("fgbl"))
	ifgshbl=CSng(rs("fgshbl"))
	ifcbl=CSng(rs("fcbl"))
	igysjxs=CSng(rs("gysjxs"))
	igysjsh=CSng(rs("gysjsh"))
	igyfcxs=CSng(rs("gyfcxs"))
	igyfcsh=CSng(rs("gyfcsh"))
	igyfgxs=CSng(rs("gyfgxs"))
	igyfgsh=CSng(rs("gyfgsh"))
	rs.close

	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	mjfz=rs("mjzf")
	gjfz=Rs("gjzf")
	ssgjf=NullToNum(Rs("ssgj"))
	qbfgjf=NullToNum(Rs("qbfgj"))
	qgjf=NullToNum(Rs("qgj"))
	hgjf=NullToNum(Rs("hgj"))
	if NullToNum(Rs("mtjgbl"))<>0 Then imtjgbl=Rs("mtjgbl")/100
	if NullToNum(Rs("dxjgbl"))<>0 Then idxjgbl=Rs("dxjgbl")/100
	select case ssgjf&qbfgjf&qgjf&hgjf
		Case "0000"			'兼容08版共挤计分模式
			'只有软硬前共挤的分值才部分加到模头部分加到定型上
			if Rs("gjfs")="3" and Rs("qhgj")="1" Then
				mtfz=Rs("mjzf")*Rs("mtbl")/100
				dxfz=Rs("mjzf")*(100-Rs("mtbl"))/100
			End if
			'软硬后共挤的分值单独加到后共挤人上
			If Rs("gjfs")="3" and Rs("qhgj")="2" Then
				mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100
				dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
			'其他情况下如果有共挤则分全加到模头
			If (not (Rs("gjfs")="3")) Then
				mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + gjfz
				dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
		Case Else		'09版共挤计分模式
			If qgjf<>0 Then
				mtgjf=qgjf*Rs("mtbl")/100
				dxgjf=qgjf-mtgjf
			End If
			mtgjf=mtgjf+ssgjf+qbfgjf
			mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + mtgjf
			dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100 + dxgjf
	end select

	bomfz=rs("bomzf")
	itsdfz=rs("tsdzf")
	sngmtbl=rs("mtbl")/100
	ijc2=datediff("d",rs("sjjssj"),rs("jhjssj"))
	ijc=0

	select case rs("rwlr")
		case "设计"
			'结构
			if not(isnull(rs("mtjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头结构',"&Round(mtfz*imtjgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtjgr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型结构',"&Round(dxfz*idxjgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxjgr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if  not(isnull(rs("gjjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','后共挤结构',"&Round(hgjf*idxjgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjjgr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			'设计
			if not(isnull(rs("mtsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头设计',"&Round(mtfz*(1-imtjgbl),1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型设计',"&Round(dxfz*(1-idxjgbl),1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','后共挤设计',"&Round(hgjf*(1-idxjgbl),1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			'在模头和定型中分别分离出结构和设计审核
			if not(isnull(rs("mtjgshr"))) then '模头结构审核
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头结构确认',"&Round(mtfz*imtjgbl*ijgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtjgshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			if not(isnull(rs("mtsjshr"))) then '模头设计审核
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				imtsjshf=Round(mtfz*(1-imtjgbl)*iljshbl,1)
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头设计确认',"&imtsjshf&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtsjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			if not(isnull(rs("dxjgshr"))) then '定型结构审核
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型结构确认',"&Round(dxfz*idxjgbl*ijgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxjgshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			if not(isnull(rs("dxsjshr"))) then '定型设计审核
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				idxsjshf=Round(dxfz*(1-idxjgbl)*iljshbl,1)
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型设计确认',"&idxsjshf&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxsjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			if not(isnull(rs("gjjgshr"))) then '后共挤结构审核
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','后共挤结构确认',"&Round(hgjf*idxjgbl*ijgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjjgshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			if not(isnull(rs("gjsjshr"))) then '后共挤设计审核
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','后共挤设计确认',"&Round(hgjf*(1-idxjgbl)*iljshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjsjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			if rs("fsr")=rs("mtsjshr") or rs("fsr")=rs("dxsjshr") or rs("fsr")=rs("gjsjshr") Then ifsxs=1
			if not(isnull(rs("fsr"))) and ifsxs=0 then '模具复审确认
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("fsr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模具复审确认',"&Round((imtsjshf+idxsjshf)*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("fsr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			'审核
			if not(isnull(rs("mtshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头审核',"&Round(mtfz*ishbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型审核',"&Round(dxfz*ishbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','后共挤审核',"&Round(hgjf*ishbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			'Bom
			if not(isnull(rs("mtbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			'调试单分值入库
			if not(isnull(rs("mttsdr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mttsdr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头调试单',"&Round(itsdfz*sngmtbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mttsdr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxtsdr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxtsdr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型调试单',"&Round(itsdfz*(1-sngmtbl),1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxtsdr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

		case "复改"
			'复改
			if not(isnull(rs("mtsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头复改',"&Round(mtfz*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型复改',"&Round(dxfz*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤复改',"&Round(gjfz*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjsjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			'审核
			if not(isnull(rs("mtshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头审核',"&Round(mtfz*ifgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型审核',"&Round(dxfz*ifgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤审核',"&Round(gjfz*ifgshbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			'BOM
			if not(isnull(rs("mtbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

			'调试单分值入库
			if not(isnull(rs("mttsdr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mttsdr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头调试单',"&Round(itsdfz*sngmtbl*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mttsdr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxtsdr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxtsdr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型调试单',"&Round(itsdfz*(1-sngmtbl)*ifgbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxtsdr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if

		case "复查"
			'复查
			if not(isnull(rs("mtshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头复查',"&Round(mtfz*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型复查',"&Round(dxfz*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤复查',"&Round(gjfz*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("gjshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			'BOM
			if not(isnull(rs("mtbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mtbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型BOM',"&Round(bomfz*0.5,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxbomr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			'调试单分值入库
			if not(isnull(rs("mttsdr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mttsdr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头调试单',"&Round(itsdfz*sngmtbl*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("mttsdr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxtsdr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxtsdr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型调试单',"&Round(itsdfz*(1-sngmtbl)*ifcbl,1)&",'"&rs("sjjssj")&"',"&ijc&",'"&rs("dxtsdr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
	end select
	rs.close
end function

Function KptoDb(kplsh,kpzrr,kpsj,kplr)
	'2017年改为延期直接考核组长,由组长在系统中考核相应组员
	dim tmpSql, tmpRs, iGroup, iKPKind, iKpMul, strbz,iPrice
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&kpzrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close
	iKpKind=3		'3为组长考评
	strBz=kplsh
	iPrice=2		'2015年开始延迟考评统一为2分/次
	iKpMul=-1
	tmpSql="select * from [kp_jsb]"
	Call xjweb.Exec("",-1)
	set tmpRs=Server.CreateObject("adodb.recordset")
	tmpRs.open tmpSql,conn,1,3
	tmpRs.AddNew
		tmpRs("kp_time")=kpsj
		tmpRs("kp_zrr")=kpzrr
		tmpRs("kp_zrrjs")="组长"
		tmpRs("kp_group")=iGroup
		tmpRs("kp_kind")=iKpKind
		tmpRs("kp_topic")="任务完成"
		tmpRs("kp_item")=kplr
		tmpRs("kp_uprice")=iPrice
		tmpRs("kp_cs")=1		'这是考评次数,系统默认为1
		tmpRs("kp_mul")=iKpMul
		tmpRs("kp_bz")=strBz
		tmpRs("kp_lsh")=kplsh
		tmpRs("kp_kpr")=session("userName")
		tmpRs.Update
	tmpRs.Close
End Function

Function ygkptodb(zrr, tt, iprice, strlsh, lsh, zrrjs)
	dim tmpSql, tmpRs, iGroup, iKPKind, strKpTopic, strKpItem, ikpmul, strbz,isjjssj
	strZrr=zrr
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

	tmpSql="select * from [mtask] where lsh='"&strlsh&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		isjjssj=tmpRs("sjjssj")
	End If
	tmpRs.Close

	iKpKind=5		'5为组员考评
	strKpTopic="任务完成"
	If tt>0 Then
		strKpItem="提前"
	Else
		strKpItem="设计延迟"
	End If
	strBz=lsh
	iPrice=2		'2015年延迟考评统一为2分/次
	If tt<0 Then
		iKpMul=-1
		tmpSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		set tmpRs=Server.CreateObject("adodb.recordset")
		tmpRs.open tmpSql,conn,1,3
		tmpRs.AddNew
			tmpRs("kp_time")=isjjssj
			tmpRs("kp_zrr")=strZrr
			tmpRs("kp_zrrjs")=zrrjs
			tmpRs("kp_group")=iGroup
			tmpRs("kp_kind")=iKpKind
			tmpRs("kp_topic")=strKpTopic
			tmpRs("kp_item")=strKpItem
			tmpRs("kp_uprice")=iPrice
			tmpRs("kp_cs")=1		'这是考评次数,系统默认为1
			tmpRs("kp_mul")=iKpMul
			tmpRs("kp_bz")=strBz
			tmpRs("kp_lsh")=strlsh
			tmpRs("kp_kpr")=session("userName")
		tmpRs.Update
		tmpRs.Close
	End If
End Function

Function Ddkp(zrr,wcsj,jgsj,lsh,zrrjs)
	'2015年修改为统一按最终结束时间考核，延迟扣2分/次，提前不考核。
	'Ddkp(责任人,完成时间,允许间隔时间,流水号,任务角色)
	dim tmpSql, tmpRs, ikp, iGroup, iPrice, iKPKind, strKpTopic, strKpItem, ikpmul, strbz, izz
	If datediff("d",wcsj,jgsj)>=0 Then
		exit Function
	else
		ikp=1
		strKpItem="设计延迟"
		iKpMul=-1
		iPrice=2
	End If
'	If (InStr(zrrjs, "审核") > 0) or (InStr(zrrjs, "确认") > 0) Then iPrice = iPrice * 0.5

	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&zrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

	iKpKind=5		'5为组员考评
	strKpTopic="任务完成"
	strBz=lsh
	tmpSql="select * from [kp_jsb]"
	Call xjweb.Exec("",-1)
	set tmpRs=Server.CreateObject("adodb.recordset")
	tmpRs.open tmpSql,conn,1,3
	tmpRs.AddNew
		tmpRs("kp_time")=wcsj
		tmpRs("kp_zrr")=zrr
		tmpRs("kp_zrrjs")=zrrjs
		tmpRs("kp_group")=iGroup
		tmpRs("kp_kind")=iKpKind
		tmpRs("kp_topic")=strKpTopic
		tmpRs("kp_item")=strKpItem
		tmpRs("kp_uprice")=iPrice
		tmpRs("kp_cs")=1		'这是考评次数,系统默认为1
		tmpRs("kp_mul")=iKpMul
		tmpRs("kp_bz")=strBz
		tmpRs("kp_lsh")=strlsh
		tmpRs("kp_kpr")=session("userName")
	tmpRs.Update
	tmpRs.Close

	'考核相应组长
	izz=""
	iPrice=1
	tmpSql="select * from [ims_user] where mid(user_able,4,1)>0 and user_group="&iGroup
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		izz=tmpRs("user_name")
	End If
	tmpRs.Close
	if izz<>"" Then
		tmpSql="select * from [kp_jsb] where kp_zrrjs like '%组长%' and Instr('提前延迟',kp_item)>0 and kp_lsh='"&strlsh&"' and kp_zrr='"&izz&"'"
		Call xjweb.Exec("",-1)
		tmpRs.open tmpSql,conn,1,3
		If tmpRs.Eof Or tmpRs.Bof Then
			tmpRs.addnew
			tmpRs("kp_time")=wcsj
			tmpRs("kp_zrr")=izz
			tmpRs("kp_zrrjs")="组长"
			tmpRs("kp_group")=iGroup
			tmpRs("kp_kind")=3
			tmpRs("kp_topic")=strKpTopic
			tmpRs("kp_item")=strKpItem
			tmpRs("kp_uprice")=iPrice
			tmpRs("kp_cs")=1		'这是考评次数,系统默认为1
			tmpRs("kp_mul")=iKpMul
			tmpRs("kp_bz")=strBz
			tmpRs("kp_lsh")=strlsh
			tmpRs("kp_kpr")=session("userName")
			tmpRs.update
		End If
		tmpRs.close
	End If
End Function

function jsdbfz(hth,khmc,rwnr,fz,jhsj,sjr,sjjs,shr,shjs)
	Dim tmpSql, tmpRs, iGroup
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&sjr&"'"
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
			rs("rwlx")="技术代表设计"
			rs("rwlr")=hth&khmc&":"&rwnr
			rs("zrr")=sjr
			rs("xz")=iGroup
			rs("zf")=fz
			rs("jssj")=sjjs
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update

	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&shr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

		rs.addnew
			rs("rwlx")="技术代表审核"
			rs("rwlr")=hth&khmc&":"&rwnr
			rs("zrr")=shr
			rs("xz")=iGroup
			rs("zf")=Round(fz/3,1)
			rs("jssj")=shjs
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update
		rs.close
		If datediff("d",sjjs,jhsj)<>0 Then Call ygkptodb(sjr, datediff("d",sjjs,jhsj), 1.5, hth, hth, "技术代表设计")
		If datediff("d",shjs,jhsj)<>0 Then Call ygkptodb(shr, datediff("d",shjs,jhsj), 0.8, hth, hth, "技术代表审核")
End Function
%>
