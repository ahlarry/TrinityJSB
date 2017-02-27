<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
	'员工考评入库文件
	Call ChkAble(0)
	Dim action, iid, dtKp, strZrr, iGroup, iKpKind, strKpTopic, strKpItem, iKpUPrice, iKpCs, iKpMul, strBz, iZlID, strkpr, strlsh, strTemp,strfz,strGroup
	action="" : iid=0 : dtKp=Now() : strZrr="" : iGroup=0 : iKpKind=0 : strKpItem="" : iKpUPrice=0 : strlsh=""
	iKpCs=1 : iKpMul=-1 : strfz=0 : strBz="" : iZlID=0 : strKpr=Session("username") : strGroup=""
	action=LCase(Request("action"))
	'数据入库函数从这里开始
	Select Case action
		Case "zrtozykp"		'3主任 to 组员考评
			iKpKind=5		'5为组员考评
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
'			If strTemp(2)="" Then
'				iKpUPrice=Request("wtfz")
'			else
'				iKpUPrice=CSng(strTemp(2))
'			End If
			iKpUPrice=Round(Request("kpfz"),1)
			If iKpUPrice="" Then Call JsAlert("考评分不能为空!","")
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")

			strlsh=Trim(Request("kplsh"))
			If strlsh<>"" Then
				strSql="Select * from [mtask] where lsh='"&strlsh&"'"
				Set Rs=xjweb.Exec(strSql,1)
				If Rs.Eof Or Rs.bof Then Call JsAlert("流水号是不是输错了?请核实一下!","")
				Dim tempzrr,mtjgshr,dxjgshr,mtsjshr,dxsjshr
				mtjgr=Rs("mtjgr") : dxjgr=Rs("dxjgr") : mtsjr=Rs("mtsjr") : dxsjr=Rs("dxsjr")
				mtjgshr=Rs("mtjgshr") : dxjgshr=Rs("dxjgshr") : mtsjshr=Rs("mtsjshr") : dxsjshr=Rs("dxsjshr")
				tempzrr=Array(mtjgr,dxjgr,mtsjr,dxsjr,mtjgshr,dxjgshr,mtsjshr,dxsjshr)
				For I = Lbound(tempzrr) to Ubound(tempzrr)
					strZrr=tempzrr(i)
					strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
					Set Rs=xjweb.Exec(strSql,1)
					If Not(Rs.Eof Or Rs.Bof) Then
						iGroup=Rs("user_group")
					Rs.Close
					iKpUPrice=CSng(strTemp(2))
					If i>3 Then	iKpUPrice=Round(iKpUPrice/3,2)
					Call kp_add("主任 → 组员")
					End If
				Next
			else
				strZrr=Request("kpzrr")
				If strZrr="" Then Call JsAlert("请选择考评人员!","")
				strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
				'response.write strsql
				Set Rs=xjweb.Exec(strSql,1)
				If Not(Rs.Eof Or Rs.Bof) Then
					iGroup=Rs("user_group")
				End If
				Rs.Close
				Call kp_add("主任 → 组员")
			End If
			If strZrr<>"" Then
				call sendmsg(strZrr, web_info(0), "考核内容:"&strKpItem&"<br>", "您因<b>"&strKpItem&"</b>被考核，详细内容请看考评列表")
			End If
			Call JsAlert("主任 → 组员考评添加成功!","ygkp_add.asp")

		Case "zrtotsykp"		'4主任 to 调试员考评
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("请选择考评人员!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=4	'4为调试员
			iKpUPrice=Round(Request("kpfz"),1)
			If iKpUPrice="" Then Call JsAlert("考评分不能为空!","")
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")
			Call kp_add("主任 → 调试员考评")
			If strZrr<>"" Then
				call sendmsg(strZrr, web_info(0), "考核内容:"&strKpItem&"<br>", "您因<b>"&strKpItem&"</b>被考核，详细内容请看考评列表")
			End If
			Call JsAlert("主任 → 调试员考评考评添加成功!","ygkp_add.asp")

		Case "zrtowgkp"		'4主任 to 网管服务人员考评
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("请选择考评人员!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=1
			iKpUPrice=Round(Request("kpfz"),1)
			If iKpUPrice="" Then Call JsAlert("考评分不能为空!","")
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")
			Call kp_add("主任 → 服务人员考评")
			If strZrr<>"" Then
				call sendmsg(strZrr, web_info(0), "考核内容:"&strKpItem&"<br>", "您因<b>"&strKpItem&"</b>被考核，详细内容请看考评列表")
			End If
			Call JsAlert("主任 → 服务人员考评添加成功!","ygkp_add.asp")

		Case "zztotsykp"		'6.组长 to 调试员考评
			dim tZrr
			tZrr=Request("kpzrr")
			If tZrr="" Then Call JsAlert("请选择考评人员!","")

			iKpKind=4		'4为调试员考评
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(strTemp(2))
			iKpMul=CInt(strTemp(3))
			strlsh=Request("hglsh")
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")

			tZrr=split(tZrr,"|")
			if Instr(strKpItem,"合格")>0 Then
				iKpUPrice=iKpUPrice/(ubound(tZrr)+1)  	'所有选择调试员平分注意加一
			else
				if ubound(tZrr)>0 Then
					'Call jsalert(ubound(tzrr) & strKpItem,"ygkp_add.asp")
					Call JsAlert("只有合格项目才能人员多选!\n\n请重新选择人员!","")
				end if
			end if

			Dim Ttsymjzf, Ttsymttsr, Ttsydxtsr
			Ttsymjzf=""	:	Ttsymttsr="" : Ttsydxtsr=""
			If Instr(strKpItem,"合格")>0 Then
				strSql="Select * from [mtask] where lsh='"&strlsh&"'"
				Set Rs=xjweb.Exec(strSql,1)
				If Rs.Eof Or Rs.bof Then Call JsAlert("流水号是不是输错了?请核实一下!","")
				Set Rs=xjweb.Exec(strSql,1)
				Ttsymjzf=Rs("mjzf")
			End If
			'二次成型
			If Instr(strKpItem,"第二次样品合格")>0 Then
			iKpUPrice=iKpUPrice*Ttsymjzf
			End If
			'三次成型
			If Instr(strKpItem,"第三次样品合格")>0 Then
			iKpUPrice=iKpUPrice*Ttsymjzf
			End If

			for i=0 to ubound(tZrr)
				strZrr=tZrr(i)
				strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"

				'response.write strsql
				Set Rs=xjweb.Exec(strSql,1)
				If Not(Rs.Eof Or Rs.Bof) Then
					iGroup=Rs("user_group")
				End If
				Rs.Close
				Call tsykp_Add()
			next
			Call JsAlert("组长 → 调试员考评成功!","ygkp_add.asp")


		Case "zztozykp"		'7.组长 to 组员考评
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("请选择考评人员!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=5		'5为组员考评
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(strTemp(2))
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")
			Call kp_add("组长 → 组员")
			Call JsAlert("组长 → 组员考评添加成功!","ygkp_add.asp")


		Case "pgbtotsykp"		'品管部 to  调试技术员考评
			Dim kpxs, strZrrjs			'考评系数,责任人角色
			dim tsZrr,tsShr, ljjs, ljxs, strbztmp
			tsZrr=Request("kpzrr")
			tsShr=Request("kpsh")
			ljjs=CInt(request("ljjs"))
			ljxs=CSng(request("ljxs"))
			iKpUPrice=CSng(Request("Pgkpfz"))
			If tsZrr="" Then Call JsAlert("请选择考评人员!","")
			if ljjs=0 Then Call JsAlert("请选择零件件数!","")
			if iKpUPrice="" Then Call JsAlert("分值不能为空!", "")

			'因为品管部一次可能涉及很多人因此在此生成随机数进行限定
			Randomize
			iZlID=rnd*99999
			iKpKind=4		'4为调试员考评
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpMul=CInt(strTemp(3))
			strbztmp=Request("kpbz") & vbcrlf & "零件件数:" & ljjs & " 系数:" & ljxs
			If Request("kpbz")="" Then Call JsAlert("备注为空!","")
			'调试审核人入库
			If tsShr<>"" Then
			strSql="Select [user_group] from [ims_user] where [user_name]='"&tsShr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close
			kpxs=1
			strZrrjs="审核"
			strZrr=tsShr
			strBz=strbztmp & vbcrlf & strZrrjs
			Call pgbkp_Add()
			End If

			'调试责任人入库
			tsZrr=split(tsZrr,"|")
			if ubound(tsZrr)>0 Then
				iKpUPrice=iKpUPrice/(ubound(tsZrr)+1)  	'所有选择调试员平分,注意加一
			end if
			for i=0 to ubound(tsZrr)
				kpxs=1
				strZrrjs="设计"
				strZrr=tsZrr(i)
				strBz=strbztmp & vbcrlf & strZrrjs
				strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
				'response.write strsql
				Set Rs=xjweb.Exec(strSql,1)
				If Not(Rs.Eof Or Rs.Bof) Then
					iGroup=Rs("user_group")
				End If
				Rs.Close
				Call pgbkp_Add()
			next
			Call JsAlert("品管部→调试技术员考评成功!","ygkp_add.asp")

		Case "pgbtozykp"		'8.品管部 to 组员考评
			strlsh=Trim(Request("kplsh"))
			If strlsh="" Then Call JsAlert("请输入相关模具的流水号!","")

			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(Request("Pgkpfz"))
			iKpMul=CInt(strTemp(3))

			'因为品管部一次可能涉及很多人因此在此生成随机数进行限定
			Randomize
			iZlID=rnd*99999

			iKpKind=5		'5为组员考评
			strBz=strlsh&","&Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")

'			If Instr(strKpItem,"样品合格")>0 Then
'				strSql="Select * from [mtask] where lsh='"&strlsh&"'"
'				Set Rs=xjweb.Exec(strSql,1)
'				If Rs.Eof Or Rs.bof Then Call JsAlert("流水号是不是输错了?请核实一下!","")
'				dim mtjgr, dxjgr, mtsjr, dxsjr, mtshr, dxshr, zmjzf,tempjs
'				mtshr=Rs("mtshr") : dxshr=Rs("dxshr") : zmjzf=Rs("mjzf")
'				mtjgr=Rs("mtjgr") : dxjgr=Rs("dxjgr") :	mtsjr=Rs("mtsjr") : dxsjr=Rs("dxsjr") :
'				mtjgshr=Rs("mtjgshr") : dxjgshr=Rs("dxjgshr") : mtsjshr=Rs("mtsjshr") : dxsjshr=Rs("dxsjshr")
'				'一次成型
'				If Instr(strKpItem,"第一次样品合格")>0 Then
'				iKpUPrice=iKpUPrice*zmjzf
'				End If
'				'二次成型
'				If Instr(strKpItem,"第二次样品合格")>0 Then
'				iKpUPrice=iKpUPrice*zmjzf
'				End If
'				tempzrr=Array(mtjgr,dxjgr,mtsjr,dxsjr,mtjgshr,dxjgshr,mtsjshr,dxsjshr,mtshr,dxshr)
'				tempjs=Array("模头结构","定型结构","模头设计","定型设计","模头结构审核","定型结构审核","模头设计审核","定型设计审核","模头审核","定型审核")
'				For I = Lbound(tempzrr) to Ubound(tempzrr)
'					strZrr=tempzrr(i)
'					If strZrr<>"" Then
'						strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
'						Set Rs=xjweb.Exec(strSql,1)
'						If Not(Rs.Eof Or Rs.Bof) Then
'							iGroup=Rs("user_group")
'						End If
'						Rs.Close
'						strZrrjs=tempjs(i)
'						If InStr(strZrrjs,"结构") Then kpxs=0.3
'						If InStr(strZrrjs,"设计") Then kpxs=0.2
'						If InStr(strZrrjs,"审核") Then kpxs=0.1
'						Call pgbkp_Add()
'					End If
'				Next
'			Else
				dim sjr, shr
				sjr="" : shr=""
				sjr=Request("kpsj")
				shr=Request("kpsh")
				ljjs=CInt(request("ljjs"))
				ljxs=CSng(request("ljxs"))
				If sjr="" Then Call JsAlert("请选择设计者!","")
				If shr="" Then Call JsAlert("请选择审核者!","")
				if ljjs=0 Then Call JsAlert("请选择零件件数!","")
				if ljxs=0.0 Then Call JsAlert("请选择零件系数!", "")
				if iKpUPrice="" Then Call JsAlert("分值不能为空!", "")

				sjr=Split(sjr,",")
				shr=Split(shr,",")
				strBz=strBz & vbcrlf & "零件件数:" & ljjs & " 系数:" & ljxs

				'sjr入库
				strGroup=""
				For i=0 to ubound(sjr)
					strSql="Select [user_group] from [ims_user] where [user_name]='"&sjr(i)&"'"
					Set Rs=xjweb.Exec(strSql,1)
					If Not(Rs.Eof Or Rs.Bof) Then
						iGroup=Rs("user_group")
					End If
					Rs.Close
					kpxs=1
					If Instr(strGroup,iGroup&",") > 0 Then
						strZrrjs="设计2"
					else
						strZrrjs="设计"
					End If
					strZrr=sjr(i)
					Call pgbkp_Add()
					strGroup=strGroup&iGroup&","
				next
				'shr入库
				strGroup=""
				For i=0 to ubound(shr)
					strSql="Select [user_group] from [ims_user] where [user_name]='"&shr(i)&"'"
					Set Rs=xjweb.Exec(strSql,1)
					If Not(Rs.Eof Or Rs.Bof) Then
						iGroup=Rs("user_group")
					End If
					Rs.Close
					kpxs=1
					If Instr(strGroup,iGroup&",") > 0 Then
						strZrrjs="审核2"
					else
						strZrrjs="审核"
					End If
					strZrr=shr(i)
					Call pgbkp_Add()
					strGroup=strGroup&iGroup&","
				Next
'			End If
			Call JsAlert("品管部 → 技术员考评添加成功!","ygkp_add.asp")

		Case "glbtozrkp"		'9. 管理部 to 主任考评
			strZrr=Request("kpzrr")
			If strZrr="" Then Call JsAlert("请选择考评人员!","")

			strSql="Select [user_group] from [ims_user] where [user_name]='"&strZrr&"'"
			'response.write strsql
			Set Rs=xjweb.Exec(strSql,1)
			If Not(Rs.Eof Or Rs.Bof) Then
				iGroup=Rs("user_group")
			End If
			Rs.Close

			iKpKind=5		'5为组员考评
			strTemp=Request("kpinfo")
			strTemp=Split(strTemp,"||")
			If Ubound(strTemp)<>3 Then Call JsAlert("请选择考评类型!","")
			strKpTopic=strTemp(0)
			strKpItem=strTemp(1)
			iKpUPrice=CSng(strTemp(2))
			iKpMul=CInt(strTemp(3))
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")
			Call kp_add("组长 → 组员")
			Call JsAlert("组长 → 组员考评添加成功!","ygkp_add.asp")

		Case "ygkpchange"		'更改考评信息
			iid=Request("id")
			If Not IsNumeric(iid) Then Call JsAlert("请从正确入口进入!","")
			iid=CLng(iid)
			strfz=Request("kpfz")
			strBz=Request("kpbz")
			If strBz="" Then Call JsAlert("备注为空!","")
			Call kp_Change()
'			Call JsAlert(strfz,"")

		Case Else
			Call JsAlert("action="&action&", 请联系管理员!","")
	End Select

	'调试员考评信息入库
	Function tsykp_Add()
		strSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("kp_time")=Now()
			Rs("kp_zrr")=strZrr
			Rs("kp_group")=iGroup
			Rs("kp_kind")=iKpKind
			Rs("kp_topic")=strKpTopic
			Rs("kp_item")=strKpItem
			Rs("kp_uprice")=iKpUPrice
			Rs("kp_cs")=1		'这是考评次数,系统默认为1
			Rs("kp_mul")=iKpMul
			If strBz<>"" Then Rs("kp_bz")=strBz
			Rs("kp_kpr")=strKpr
		Rs.Update
		Rs.Close
	End Function

	'考评信息入库
	Function kp_Add(str)
		strSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("kp_time")=Now()
			Rs("kp_zrr")=strZrr
			Rs("kp_group")=iGroup
			Rs("kp_kind")=iKpKind
			Rs("kp_topic")=strKpTopic
			Rs("kp_item")=strKpItem
			Rs("kp_uprice")=iKpUPrice
			Rs("kp_cs")=1		'这是考评次数,系统默认为1
			Rs("kp_mul")=iKpMul
			If strBz<>"" Then Rs("kp_bz")=strBz
			Rs("kp_kpr")=strKpr
		Rs.Update
		Rs.Close
	End Function

	'品管部考评信息入库
	Function pgbkp_Add()
		Dim strkptime
		strkptime=Request("khsj")
		If strkptime="" Then strkptime=Now()
		strSql="select * from [kp_jsb]"
		Call xjweb.Exec("",-1)
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("kp_time")=strkptime
			Rs("kp_zrr")=strZrr
			Rs("kp_zrrjs")=strZrrjs
			Rs("kp_group")=iGroup
			Rs("kp_kind")=iKpKind
			Rs("kp_topic")=strKpTopic
			Rs("kp_item")=strKpItem
			Rs("kp_uprice")=iKpUPrice * kpxs
			Rs("kp_cs")=1		'这是考评次数,系统默认为1
			Rs("kp_mul")=iKpMul
			If strlsh<>"" Then Rs("kp_lsh")=strlsh
			Rs("kp_zlid")=iZlID
			If strBz<>"" Then Rs("kp_bz")=strBz
			Rs("kp_kpr")=strKpr
		Rs.Update
		Rs.Close
	End Function

	'更改考评信息入库
	Function kp_Change()
	Dim strFeedBack, strgzz, strkpjs, strclsh, striPage, strkptime, strkpitem
	strkptime=Request("kpsj")
	strZrr=Trim(Request("zrr"))
	strkpitem = trim(request("kpitem"))
	strkpjs = trim(request("kpjs"))
	strclsh = trim(request("kplsh"))
	strgzz =request("kpgzz")
	striPage =request("ipage")
	strFeedBack=""
	If strZrr<>"" Then strFeedBack="zrr="&strZrr
	If strkpitem<>"" Then strFeedBack="kpitem="&strkpitem&"&"&strFeedBack
	If strgzz<>"" Then strFeedBack="gzz="&strgzz&"&"&strFeedBack
	If strkpjs<>"" Then strFeedBack="kpjs="&strkpjs&"&"&strFeedBack
	If strclsh<>"" Then strFeedBack="clsh="&strclsh&"&"&strFeedBack
	If striPage<>"0" Then strFeedBack="iPage="&striPage&"&"&strFeedBack
	If strFeedBack<>"" Then strFeedBack="?"&strFeedBack

		'检测ID号是否存在
		Set Rs=xjweb.Exec("select * from [kp_jsb] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("此技术考评信息可能已经删除!","ygkp_list.asp"&strFeedBack)
			Rs.Close
			Exit Function
		End If
		Rs.Close
			strSql="select * from [kp_jsb] where id=" & iid
			Call xjweb.Exec("",-1)
			Rs.open strSql,conn,1,3
				If strBz<>"" Then Rs("kp_bz")=strBz
				If strfz<>"" Then Rs("kp_uprice")=strfz
				If strkptime<>"" Then Rs("kp_time")=strkptime
			Rs.update
			Rs.close

		Call JsAlert("员工考评更改成功","ygkp_list.asp"&strFeedBack)
	End Function
%>
