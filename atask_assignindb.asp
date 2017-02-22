<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble("4,6")
	'本文件只负责分配辅助任务的入库
'	Call JsAlert("用户名或密码不正确,请核实后再输!","")
	Dim strfplr, strzrr, strlsh
	strfplr=Request("fplr")
	strzrr=Request("zrr")
	strlsh=Request("lsh")
	If strfplr="" Or (strzrr="" And InStr(strfplr,"开始")>0) Or strlsh="" Then
		Call JsAlert("分配辅助任务信息不够","atast_assign.asp")
	End If
	strSql="select * from [mtask] where lsh='"&strlsh&"'"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("指定流水号的任务书不存在,请核实!","atask_assign.asp")
	End If
	Rs.Close

	Dim mtbl, tsdzf, tszf, tsxxzlzf, ZstrSql		'模头比例, 调试单总分, 调试总分, 调试信息整理总分
	ZstrSql="select * from [mtask] where lsh='"&strlsh&"'"
	Set Rs=xjweb.Exec(ZstrSql, 1)
		mtbl=Rs("mtbl")
		tsdzf=Rs("tsdzf")
		tszf=Rs("tszf")
		tsxxzlzf=Rs("tsxxzlzf")

	strSql=""
	Select case strfplr
		case "开始模头调试单"
			strSql="update [mtask] set mttsdr='"&strzrr&"', mttsdks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "开始定型调试单"
			strSql="update [mtask] set dxtsdr='"&strzrr&"', dxtsdks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "开始全套调试单"
			strSql="update [mtask] set mttsdr='"&strzrr&"', mttsdks='"&now()&"', dxtsdr='"&strzrr&"', dxtsdks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)

		case "开始模头调试"
			strSql="update [mtask] set mttsr='"&strzrr&"', mttsks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
			call sendmsg(Rs("mtjgr"), web_info(0), "模头调试开始", "流水号 <b>"&strlsh&"</b> 开始模头调试</a>")

			If xjweb.RsCount("[ts_mould] where lsh='"&strlsh&"'")=0 Then
				strSql="insert into ts_mould (lsh,tskssj) values ('"&strlsh&"','"&now()&"')"
				call xjweb.Exec(strSql, 0)
			End If
		case "开始定型调试"
			strSql="update [mtask] set dxtsr='"&strzrr&"', dxtsks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
			call sendmsg(Rs("dxjgr"), web_info(0), "定型调试开始", "流水号 <b>"&strlsh&"</b> 开始定型调试</a>")

			If xjweb.RsCount("[ts_mould] where lsh='"&strlsh&"'")=0 Then
				strSql="insert into ts_mould (lsh,tskssj) values ('"&strlsh&"','"&now()&"')"
				call xjweb.Exec(strSql, 0)
			End If

		case "开始全套调试"
			strSql="update [mtask] set mttsr='"&strzrr&"', mttsks='"&now()&"', dxtsr='"&strzrr&"', dxtsks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
			call sendmsg(Rs("mtjgr"), web_info(0), "全套调试开始", "流水号 <b>"&strlsh&"</b> 开始全套调试</a>")
			call sendmsg(Rs("dxjgr"), web_info(0), "全套调试开始", "流水号 <b>"&strlsh&"</b> 开始全套调试</a>")

			strSql="insert into ts_mould (lsh,tskssj) values ('"&strlsh&"','"&now()&"')"
			call xjweb.Exec(strSql, 0)

		case "开始模头调试信息整理"
			strSql="update [mtask] set mttsxxzlr='"&strzrr&"', mttsxxzlks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "开始定型调试信息整理"
			strSql="update [mtask] set dxtsxxzlr='"&strzrr&"', dxtsxxzlks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "开始全套调试信息整理"
			strSql="update [mtask] set mttsxxzlr='"&strzrr&"', mttsxxzlks='"&now()&"', dxtsxxzlr='"&strzrr&"', dxtsxxzlks='"&now()&"' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)

		case "结束模头调试单"
			if not bataskend("mttsdjs", strlsh) then
				strSql="update [mtask] set mttsdjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "mttsd")
			end if

		case "结束定型调试单"
			if not bataskend("dxtsdjs", strlsh) then
				strSql="update [mtask] set dxtsdjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "dxtsd")
			end if

		case "结束全套调试单"
			if not bataskend("mttsdjs", strlsh) and not bataskend("dxtsdjs", strlsh) then
				strSql="update [mtask] set mttsdjs='"&now()&"', dxtsdjs='"&now()&"'  where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "qttsd")
			end if

		case "结束模头调试"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "mtts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "模头调试结束", "流水号 <b>"&strlsh&"</b> 结束模头调试</a>")
			end if
		case "结束定型调试"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'确认模具调试是否结束
				call fentodb(strlsh, "dxts")
				call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "定型调试结束", "流水号 <b>"&strlsh&"</b> 结束定型调试</a>")
			end if
		case "结束全套调试"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "qtts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "全套调试结束", "流水号 <b>"&strlsh&"</b> 结束全套调试</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "全套调试结束", "流水号 <b>"&strlsh&"</b> 结束全套调试</a>")
			end if
		case "模头厂内初调"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "mtcts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "模头厂内初调", "流水号 <b>"&strlsh&"</b> 模头厂内初调</a>")
			end if
		case "定型厂内初调"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'确认模具调试是否结束
				call fentodb(strlsh, "dxcts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "定型厂内初调", "流水号 <b>"&strlsh&"</b> 定型厂内初调</a>")
			end if
		case "全套厂内初调"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "qtcts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "全套厂内初调", "流水号 <b>"&strlsh&"</b> 全套厂内初调</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "全套厂内初调", "流水号 <b>"&strlsh&"</b> 全套厂内初调</a>")
			end if
		case "模头厂外精调"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "mtjts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "模头厂外精调", "流水号 <b>"&strlsh&"</b> 模头厂外精调</a>")
			end if
		case "定型厂外精调"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'确认模具调试是否结束
				call fentodb(strlsh, "dxjts")
				call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "定型厂外精调", "流水号 <b>"&strlsh&"</b> 定型厂外精调</a>")
			end if
		case "全套厂外精调"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "qtjts")
				call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "全套厂外精调", "流水号 <b>"&strlsh&"</b> 全套厂外精调</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "全套厂外精调", "流水号 <b>"&strlsh&"</b> 全套厂外精调</a>")
			end if
		case "模头预验收或寄样"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "mtjyts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "模头预验收或寄样", "流水号 <b>"&strlsh&"</b> 模头预验收或寄样</a>")
			end if
		case "定型预验收或寄样"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'确认模具调试是否结束
				call fentodb(strlsh, "dxjyts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "定型预验收或寄样", "流水号 <b>"&strlsh&"</b> 定型预验收或寄样</a>")
			end if
		case "全套预验收或寄样"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "qtjyts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "全套预验收或寄样", "流水号 <b>"&strlsh&"</b> 全套预验收或寄样</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "全套预验收或寄样", "流水号 <b>"&strlsh&"</b> 全套预验收或寄样</a>")
			end if
		case "模头来厂验收"
			if not bataskend("mttsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "mtysts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "模头来厂验收", "流水号 <b>"&strlsh&"</b> 模头来厂验收</a>")
			end if
		case "定型来厂验收"
			if not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)		'确认模具调试是否结束
				call fentodb(strlsh, "dxysts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("dxjgr"), web_info(0), "定型来厂验收", "流水号 <b>"&strlsh&"</b> 定型来厂验收</a>")
			end if
		case "全套来厂验收"
			if not bataskend("mttsjs", strlsh) and not bataskend("dxtsjs", strlsh) then
				strSql="update [mtask] set mttsjs='"&now()&"', dxtsjs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call check_tsend(strlsh)	'确认模具调试是否结束
				call fentodb(strlsh, "qtysts")
			'	call TscsKp(strlsh)
				call sendmsg(Rs("mtjgr"), web_info(0), "全套来厂验收", "流水号 <b>"&strlsh&"</b> 全套来厂验收</a>")
				call sendmsg(Rs("dxjgr"), web_info(0), "全套来厂验收", "流水号 <b>"&strlsh&"</b> 全套来厂验收</a>")
			end if
		case "结束模头调试信息整理"
			if not bataskend("mttsxxzljs", strlsh) then
				strSql="update [mtask] set mttsxxzljs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "mttsxxzl")
			end if
		case "结束定型调试信息整理"
			if not bataskend("dxtsxxzljs", strlsh) then
				strSql="update [mtask] set dxtsxxzljs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				call fentodb(strlsh, "dxtsxxzl")
			end if
		case "结束全套调试信息整理"
			if not bataskend("mttsxxzljs", strlsh) and not bataskend("dxtsxxzljs", strlsh) then
				strSql="update [mtask] set mttsxxzljs='"&now()&"', dxtsxxzljs='"&now()&"' where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				'call fentodb(strlsh, "qttsxxzl")
				call fentodb(strlsh, "mttsxxzl")
				call fentodb(strlsh, "dxtsxxzl")
			end if

		case else
			Call JsAlert("(系统异常)请联系管理员!","")
	end select
	Rs.Close

	'判断是否全部完成! 以及是否可以调试
	dim bok
	bok=false
	strSql="select * from [mtask] where lsh='"&strlsh&"'"
	set rs=xjweb.Exec(strSql, 1)
	select case rs("mjxx")
		case "全套"
			if not(isnull(rs("mttsxxzljs"))) and not(isnull(rs("dxtsxxzljs"))) then
				strSql="update [mtask] set mjjs=true where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				bok=true
			end if
		case "模头"
			if not(isnull(rs("mttsxxzljs"))) then
				strSql="update [mtask] set mjjs=true where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				bok=true
			end if
		case "定型"
			if not(isnull(rs("dxtsxxzljs"))) then
				strSql="update [mtask] set mjjs=true where lsh='"&strlsh&"'"
				call xjweb.Exec(strSql,0)
				bok=true
			end if
	end select
	rs.close

	If bok Then
		Call JsAlert("流水号 【"&strlsh&"】 任务全部完成!","atask_assign.asp")
	Else
		Call JsAlert("流水号 【"&strlsh&"】 调试任务分配 ［"&strfplr&"］ 成功!", "atask_assign.asp?s_lsh="&strlsh&"")

	End If


Rem 结束
function bataskend(trs, lsh)			'判断任务是否已经完成,防止进行两次记分
	bataskend=true
	strSql="select "&trs&" from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	if isnull(rs(trs)) then bataskend=false
	rs.close
	if bataskend then
		response.write("<script language=""javascript"">alert('此任务已经结束!');location.href='atask_assign.asp';</script>")
		response.end
	end if
end function

function fentodb(lsh, strlr)
	'将分值写入分值库
	dim itsxs, itslb, iedsx, iedxx, itscs, itsdfz, itsllfz, itsfz, itsxxzlfz, sngmtbl, strzrr, dtjssj, sjjssj, jhsj
	dim ijt, ifgbl, ifcbl, irwnr
	ijt=0

	strSql="select * from [c_fzbl]"
	set rs=xjweb.Exec(strSql, 1)
	ifgbl=CSng(rs("fgbl"))
	ifcbl=CSng(rs("fcbl"))
	rs.close

	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql, 1)
	irwnr=rs("rwlr")
	jhsj=rs("jhjssj")
	itsdfz=rs("tsdzf")
	itslb=rs("tslb")
	itsllfz=rs("tszf")
	itsxxzlfz=rs("tsxxzlzf")
	sngmtbl=rs("mtbl") / 100
	Rs.Close

	select case irwnr
		case "设计"
			irwnr=""
			ifgbl=1
		case "复查"
			ifgbl=ifcbl
	end select

	'厂外精调
	If Right(strlr,3)="jts" Then
		strSql="select * from [ts_tsxx] where lsh='"&lsh&"' and tslr like '%精调%'"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		If Rs.Eof Or Rs.Bof Then
			rs.addnew
				rs("lsh")=lsh
				rs("tsyy")="厂外精调"
				rs("tslr")="厂外精调"
				rs("tssj")=now()
				rs("tsr")="Sj901"
			rs.update
			ijt=1
		End If
		rs.close

		If ijt=1 Then
			strSql="select * from [ts_mould] where lsh='"&lsh&"'"
			Call xjweb.Exec("",-1)
			rs.open strSql,conn,1,3
				if isnull(rs("tskssj")) then rs("tskssj")=now()
				rs("tscs")=rs("tscs") + 1
				rs("tsgxsj")=now()
			rs.update
			rs.close
		End If
	End If
	'预验收或寄样、来厂验收、厂内初调
	If InStr(strlr,"jyts")>0 Then itsllfz=itsllfz*1.5
	If InStr(strlr,"ysts")>0 Then itsllfz=itsllfz*2
	If InStr(strlr,"cts")>0 Then itsllfz=itsllfz*0.75
	'================
	itsxs=1
	If not(IsNull(itslb)) and itslb<>"A类" and right(strlr,2)="ts" and right(strlr,3)<>"cts" Then
		strSql="select * from [c_tscs] where dmlb='"&itslb&"'"
		set rs=xjweb.Exec(strSql, 1)
			If not(Rs.Eof Or Rs.Bof) Then
				iedsx=rs("edsx")
				iedxx=rs("edxx")
			else
				iedsx=0
			End If
		Rs.Close
		If iedsx<>0 Then
			strSql="select * from [ts_mould] where LSH='"&lsh&"'"
			set rs=xjweb.Exec(strSql, 1)
			itscs=rs("TSCS")
			Rs.Close
			If iedxx > itscs Then itsxs=1+(iedxx-itscs)*0.15
			If iedsx < itscs Then itsxs=1+(iedsx-itscs)*0.15
			If itsxs < 0.6 Then itsxs=0.6
		End If
	End If
	'预验收或寄样、来厂验收只加分不减分
	If InStr(strlr,"jyts")>0 or InStr(strlr,"ysts")>0 Then
		If itsxs<1 Then
			itsxs=1
		End If
	End If
	'实际调试分值
	itsfz=Round(itsllfz*itsxs,1)

	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	sjjssj=rs("sjjssj")
	select case strlr
		case "mttsd"	'模头调试单
			strzrr=rs("mttsdr")
			dtjssj=rs("mttsdjs")
'			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','模头调试单',"&itsdfz*sngmtbl*ifgbl&",'"&dtjssj&"',0,'"&strzrr&"')"
'			call xjweb.Exec(strSql,0)
			If (datediff("d", dtjssj, jhsj) < -20) and (not isnull(rs("mttsdjs"))) and (not isnull(rs("dxtsdjs"))) then
				call Tsdkp(strzrr, strlsh, 5, strlsh)
			End if
		case "dxtsd"	'定型调试单
			strzrr=rs("dxtsdr")
			dtjssj=rs("dxtsdjs")
'			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','定型调试单',"&itsdfz*(1-sngmtbl)*ifgbl&",'"&dtjssj&"',0,'"&strzrr&"')"
'			call xjweb.Exec(strSql,0)
			If (datediff("d", dtjssj, jhsj) < -20) and (not isnull(rs("mttsdjs"))) and (not isnull(rs("dxtsdjs"))) then
				call Tsdkp(strzrr, strlsh, 5, strlsh)
			End if
		case "qttsd"	'全套调试单
			strzrr=rs("mttsdr")
			dtjssj=rs("mttsdjs")
'			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','全套调试单',"&itsdfz*ifgbl&",'"&dtjssj&"',0,'"&strzrr&"')"
'			call xjweb.Exec(strSql,0)
			If (datediff("d", dtjssj, jhsj) < -20) and (not isnull(rs("mttsdjs"))) and (not isnull(rs("dxtsdjs"))) then
				call Tsdkp(strzrr, strlsh, 5, strlsh)
			End if
		case "mtts"		'模头调试
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','模头调试合格',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='调试合格' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxts"		'定型调试
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','定型调试合格',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='调试合格' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtts"		'全套调试
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','全套调试合格',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='调试合格' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtcts"		'模头厂内初调
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','模头厂内初调',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='厂内初调' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxcts"		'定型厂内初调
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','定型厂内初调',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='厂内初调' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtcts"		'全套厂内初调
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','全套厂内初调',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='厂内初调' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtjts"		'模头厂外精调
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','模头厂外精调',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='厂外精调' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxjts"		'定型厂外精调
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','定型厂外精调',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='厂外精调' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtjts"		'全套厂外精调
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','全套厂外精调',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='厂外精调' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtjyts"		'模头预验收或寄样
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','模头预验收或寄样',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='预验收或寄样' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxjyts"		'定型预验收或寄样
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','定型预验收或寄样',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='预验收或寄样' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtjyts"		'全套预验收或寄样
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','全套预验收或寄样',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='预验收或寄样' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mtysts"		'模头来厂验收
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','模头来厂验收',"&itsllfz*sngmtbl&","&itsfz*sngmtbl&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='来厂验收' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "dxysts"		'定型来厂验收
			strzrr=rs("dxtsr")
			dtjssj=rs("dxtsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','定型来厂验收',"&itsllfz*(1-sngmtbl)&","&itsfz*(1-sngmtbl)&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='来厂验收' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "qtysts"		'全套来厂验收
			strzrr=rs("mttsr")
			dtjssj=rs("mttsjs")
			strSql="insert into [mantime] (lsh, rwlr, rwfz, fz, jssj, jc, zrr) values ('"&lsh&"','全套来厂验收',"&itsllfz&","&itsfz&",'"&dtjssj&"','"&itsxs&"','"&strzrr&"')"
			call xjweb.Exec(strSql,0)
			strSql="update [ts_mould] set tsjg='来厂验收' where lsh='"&strlsh&"'"
			call xjweb.Exec(strSql,0)
		case "mttsxxzl"		'模头调试信息整理
			strzrr=rs("mttsxxzlr")
			dtjssj=rs("mttsxxzljs")
			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','模头调试信息整理',"&itsxxzlfz*sngmtbl&",'"&dtjssj&"',0,'"&strzrr&"')"
			call xjweb.Exec(strSql,0)
		case "dxtsxxzl"		'定型调试信息整理
			strzrr=rs("dxtsxxzlr")
			dtjssj=rs("dxtsxxzljs")
			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','定型调试信息整理',"&itsxxzlfz*(1-sngmtbl)&",'"&dtjssj&"',0,'"&strzrr&"')"
			call xjweb.Exec(strSql,0)
		case "qttsxxzl"		'全套调试信息整理
			strzrr=rs("mttsxxzlr")
			dtjssj=rs("mttsxxzljs")
			strSql="insert into [mantime] (lsh, rwlr, fz, jssj, jc, zrr) values ('"&lsh&"','全套调试信息整理',"&itsxxzlfz&",'"&dtjssj&"',0,'"&strzrr&"')"
			call xjweb.Exec(strSql,0)
		case else
	end select
end function

function check_tsend(strlsh)
	strSql="select * from [mtask] where lsh='"&strlsh&"'"
	set rs=xjweb.Exec(strSql, 1)
	if not(rs.eof or rs.bof) then
		select case rs("mjxx")
			case "全套"
				if not(isnull(rs("mttsjs"))) and not(isnull(rs("dxtsjs"))) then
					strSql="update [ts_mould] set tsjssj='"&now()&"' where lsh='"&strlsh&"'"
					call xjweb.Exec(strSql, 0)
				end if
			case "模头"
				if not(isnull(rs("mttsjs"))) then
					strSql="update [ts_mould] set tsjssj='"&now()&"' where lsh='"&strlsh&"'"
					call xjweb.Exec(strSql, 0)
				end if
			case "定型"
				if not(isnull(rs("dxtsjs"))) then
					strSql="update [ts_mould] set tsjssj='"&now()&"' where lsh='"&strlsh&"'"
					call xjweb.Exec(strSql, 0)
				end if
		end select
	end if
	rs.close
end function

'Function Tsdkp(lsh)	'当调试单延期时取消模具的提前分,暂不执行
'	Dim tmpRs
'	tmpRs="delete from [kp_jsb] where kp_lsh='"&lsh&"'and kp_item='提前'"
'	call xjweb.Exec(tmpRs, 0)
'End Function

Function Tsdkp(strzrr, lsh, iprice, bz)
	dim tmpSql, tmpRs, iGroup, iKPKind, strKpTopic, strKpItem, ikpmul, strbz, strSql
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strzrr&"'"
	Set tmpRs=xjweb.Exec(tmpSql,1)
	If Not(tmpRs.Eof Or tmpRs.Bof) Then
		iGroup=tmpRs("user_group")
	Else
		iGroup=0
	End If
	tmpRs.Close

	iKpKind=5		'5为组员考评
	strKpTopic="任务完成"
		strKpItem="延迟"
		iKpMul=-1
		strBz=bz
		tmpSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		set tmpRs=Server.CreateObject("adodb.recordset")
		tmpRs.open tmpSql,conn,1,3
		tmpRs.AddNew
			tmpRs("kp_time")=Now()
			tmpRs("kp_zrr")=strZrr
			tmpRs("kp_zrrjs")="调试手册"
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
End Function

Function TscsKp(slsh)	'根据调试次数对技术员及组长进行考评
	Dim KpSql, KpRs, sngmtbl, iddh, idmmc, ijgzz, ijsdb, igfwh, imtjgr, imtjgshr, idxjgr, idxjgshr, imtsjr, imtsjshr, idxsjr, idxsjshr, imjzf, itslb, iedsx, iedxx, itscs, itsxs, ifz, iGroup, strKpItem, iKpMul, kp_bz, iZlID
	'获得随机数，用于分辨是否同一批入库数据
	Randomize
	iZlID=rnd*99999
	'获取本套模具实际调试次数
	KpSql="select * from [ts_mould] where lsh='"&slsh&"'"
	set KpRs=xjweb.Exec(KpSql, 1)
	if not(KpRs.eof or KpRs.bof) then
		If isnull(KpRs("tsjssj")) then
			KpRs.close
			exit Function
		else
			itscs=KpRs("tscs")
		End If
	end if
	KpRs.close
	'获取模具调试类型、设计人等信息
	KpSql="select * from [mtask] where lsh='"&slsh&"'"
	set KpRs=xjweb.Exec(KpSql, 1)
		ijgzz=KpRs("jgzz")						'组长
		imtjgr=KpRs("mtjgr")					'模头结构人
		imtjgshr=KpRs("mtjgshr")				'模头结构审核人
		idxjgr=KpRs("dxjgr")					'定型结构人
		idxjgshr=KpRs("dxjgshr")				'定型结构审核人
		itslb=KpRs("tslb")						'调试类别
	KpRs.Close
	'获取本套模具额定调试次数
	If isnull(itslb) Then exit Function
	KpSql="select * from [c_tscs] where dmlb='"&itslb&"'"
	set KpRs=xjweb.Exec(KpSql, 1)
		If not(KpRs.Eof Or KpRs.Bof) Then
			iedsx=KpRs("edsx")
			iedxx=KpRs("edxx")
		else
			KpRs.Close
			exit Function
		End If
	KpRs.Close
	'计算超出标准应考核调试次数
	itsxs=0
	If iedxx > itscs Then itsxs=iedxx-itscs
	If iedsx < itscs Then itsxs=iedsx-itscs
	If itsxs=0 or (Instr("B类C类",itslb)>0 and itsxs<0) Then
		 exit Function
	End If
	'=====================
	'对结构设计人进行考核
'	If not(isNull(imtjgr)) Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&imtjgr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="厂内调试少于额定次数"
'		else
'			strKpItem="厂内调试多于额定次数"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=imtjgr
'			KpRs("kp_zrrjs")="模头结构"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="设计质量"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"额定:"&iedxx&"-"&iedsx&"次,实际:"&itscs&"次"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
'
'	If not(isNull(idxjgr)) Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&idxjgr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="厂内调试少于额定次数"
'		else
'			strKpItem="厂内调试多于额定次数"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=idxjgr
'			KpRs("kp_zrrjs")="定型结构"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="设计质量"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"额定:"&iedxx&"-"&iedsx&"次,实际:"&itscs&"次"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
'	'对结构审核人进行考核
'	If not(isNull(imtjgshr)) and imtjgshr<>ijgzz Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&imtjgshr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="厂内调试少于额定次数"
'		else
'			strKpItem="厂内调试多于额定次数"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=imtjgshr
'			KpRs("kp_zrrjs")="模头结构审核"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="设计质量"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"额定:"&iedxx&"-"&iedsx&"次,实际:"&itscs&"次"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
'	If not(isNull(idxjgshr)) and idxjgshr<>ijgzz Then
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&idxjgshr&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="厂内调试少于额定次数"
'		else
'			strKpItem="厂内调试多于额定次数"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=idxjgshr
'			KpRs("kp_zrrjs")="定型结构审核"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=5
'			KpRs("kp_topic")="设计质量"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"额定:"&iedxx&"-"&iedsx&"次,实际:"&itscs&"次"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If
	'对组长进行考核(结构组长(ijgzz)、调试组长(tszz))
'	Dim xxzz, tszz
'	KpSql="Select * from [ims_user] where mid(user_able,4,1)>0"
'	Set KpRs=xjweb.Exec(KpSql,1)
'		Do While Not KpRs.Eof
'			If KpRs("user_group")=5 and mid(KpRs("user_able"),6,1)>0 and mid(KpRs("user_able"),3,1)=0 Then tszz=KpRs("user_name")
'			KpRs.moveNext
'		Loop
'	KpRs.Close

'	If not(isNull(ijgzz)) Then
'		ifz=Round(imjzf*sngmtbl*0.1*itsxs,1)
'		KpSql="Select [user_group] from [ims_user] where [user_name]='"&ijgzz&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'		KpRs.Close
'
'		If iedxx > itscs Then
'			iKpMul=1
'		else
'			iKpMul=-1
'		End If
'		If iKpMul=1 Then
'			strKpItem="厂内调试少于额定次数"
'		else
'			strKpItem="厂内调试多于额定次数"
'		End If
'		KpSql="select * from [kp_jsb]"
'		Call xjweb.Exec("",-1)
'		KpRs.open KpSql,conn,1,3
'		KpRs.AddNew
'			KpRs("kp_time")=Now()
'			KpRs("kp_zrr")=ijgzz
'			KpRs("kp_zrrjs")="结构组长"
'			KpRs("kp_group")=iGroup
'			KpRs("kp_kind")=3
'			KpRs("kp_topic")="设计质量"
'			KpRs("kp_item")=strKpItem
'			KpRs("kp_uprice")=itsxs*0.4
'			KpRs("kp_mul")=iKpMul
'			KpRs("kp_bz")=slsh&"额定:"&iedxx&"-"&iedsx&"次,实际:"&itscs&"次"
'			KpRs("kp_kpr")="Sj901"
'			KpRs("kp_lsh")=slsh
'			KpRs("kp_zlid")=iZlID
'		KpRs.Update
'		KpRs.Close
'	End If

'	KpSql="Select [user_group] from [ims_user] where [user_name]='"&tszz&"'"
'		Set KpRs=xjweb.Exec(KpSql,1)
'		If Not(KpRs.Eof Or KpRs.Bof) Then
'			iGroup=KpRs("user_group")
'		End If
'	KpRs.Close
'	KpSql="select * from [kp_jsb]"
'	Call xjweb.Exec("",-1)
'	KpRs.open KpSql,conn,1,3
'	KpRs.AddNew
'		KpRs("kp_time")=Now()
'		KpRs("kp_zrr")=tszz
'		KpRs("kp_zrrjs")="调试组长"
'		KpRs("kp_group")=iGroup
'		KpRs("kp_kind")=3
'		KpRs("kp_topic")="设计验证"
'		KpRs("kp_item")=strKpItem
'		KpRs("kp_uprice")=itsxs*0.2
'		KpRs("kp_mul")=iKpMul
'		KpRs("kp_bz")=slsh&"额定:"&iedxx&"-"&iedsx&"次,实际:"&itscs&"次"
'		KpRs("kp_kpr")="Sj901"
'		KpRs("kp_lsh")=slsh
'		KpRs("kp_zlid")=iZlID
'		KpRs.Update
'	KpRs.Close

	set KpRs = nothing
End Function
%>