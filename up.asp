<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="升级数据"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage=""
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, strlsh
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd height=300>
      <%
      strlsh=split("11221,11249,11250,11251,11252,11253,11254,11256,11257,11258",",")
      for i=0 to ubound(strlsh)
     	 Call FenToDB(strlsh(i))
      next
      %>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function Updata()
	Dim mystr, mystr1, ikhxs, rwlr_change, tmpRs
	ikhxs=1
	strSql="select * from [ftask] where datediff('m',jssj,'"&now()&"')<4 and (rwlx='零星修理' or rwlx='零星任务' or rwlx='技术代表设计')"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Do while not Rs.eof
				if IsNull(Rs("ed")) Then Rs("ed")=0
				Rs.update
		Rs.movenext
		Loop
	Rs.close
end function

Function FenToDB(lsh)
	'将分值写入分值库
	Dim mjfz, mtfz, dxfz, gjfz, bomfz, ijgbl, isjbl, ishbl, ifgbl, ifgshbl, ifcbl, ijc, ijc2, imtjgbl, idxjgbl, ijgshbl, iljshbl, iwcsj, mtgjf, dxgjf, ssgjf, qbfgjf, qgjf, hgjf
	Dim igysjxs, igysjsh, igyfcxs, igyfcsh, igyfgxs, igyfgsh, iGroup, tmpSql, tmpRs
	mtgjf=0 : dxgjf=0
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
	ijc2=datediff("d",now(),rs("jhjssj"))
	ijc=0

	Dim ijgsj, isj
	If IsNull(rs("jhjgsj")) Then
		isj=INT(datediff("d", rs("jhkssj"), rs("jhjssj"))/2)
		ijgsj=dateadd("d",isj,rs("jhkssj"))
	else
		ijgsj=rs("jhjgsj")
	End if
'考核中模头和定型结构结束时间取两者最晚的一个
	select case rs("rwlr")
		case "设计"
			'工艺
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头工艺设计',"&Round(mtfz*igysjxs,1)&",'"&now()&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型工艺设计',"&Round(dxfz*igysjxs,1)&",'"&now()&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤工艺设计',"&Round(hgjf*igysjxs,1)&",'"&now()&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("mtgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头工艺审核',"&Round(mtfz*igysjsh,1)&",'"&now()&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型工艺审核',"&Round(dxfz*igysjsh,1)&",'"&now()&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤工艺审核',"&Round(hgjf*igysjsh,1)&",'"&now()&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
		case "复改"
			'工艺
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头工艺复改',"&Round(mtfz*igyfgxs,1)&",'"&now()&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型工艺复改',"&Round(dxfz*igyfgxs,1)&",'"&now()&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤工艺复改',"&Round(hgjf*igyfgxs,1)&",'"&now()&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("mtgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头工艺复改审核',"&Round(mtfz*igyfgsh,1)&",'"&now()&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型工艺复改审核',"&Round(dxfz*igyfgsh,1)&",'"&now()&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤工艺复改审核',"&Round(hgjf*igyfgsh,1)&",'"&now()&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
		case "复查"
			'工艺
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头工艺复查',"&Round(mtfz*igyfcxs,1)&",'"&now()&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型工艺复查',"&Round(dxfz*igyfcxs,1)&",'"&now()&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤工艺复查',"&Round(hgjf*igyfcxs,1)&",'"&now()&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("mtgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','模头工艺复查审核',"&Round(mtfz*igyfcsh,1)&",'"&now()&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("dxgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','定型工艺复查审核',"&Round(dxfz*igyfcsh,1)&",'"&now()&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isNull(rs("gjgyshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjgyshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','共挤工艺复查审核',"&Round(hgjf*igyfcsh,1)&",'"&now()&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
	end select
	rs.close
end function
%>
