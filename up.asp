<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="��������"					'ҳ�������λ��( ��������� �� ���������)
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
	strSql="select * from [ftask] where datediff('m',jssj,'"&now()&"')<4 and (rwlx='��������' or rwlx='��������' or rwlx='�����������')"
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
	'����ֵд���ֵ��
	Dim mjfz, mtfz, dxfz, gjfz, bomfz, ijgbl, isjbl, ishbl, ifgbl, ifgshbl, ifcbl, ijc, ijc2, imtjgbl, idxjgbl, ijgshbl, iljshbl, iwcsj, mtgjf, dxgjf, ssgjf, qbfgjf, qgjf, hgjf
	Dim igysjxs, igysjsh, igyfcxs, igyfcsh, igyfgxs, igyfgsh, iGroup, tmpSql, tmpRs
	mtgjf=0 : dxgjf=0
	'ijc===���ͷ�ֵ
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
		Case "0000"			'����08�湲���Ʒ�ģʽ
			'ֻ����Ӳǰ�����ķ�ֵ�Ų��ּӵ�ģͷ���ּӵ�������
			if Rs("gjfs")="3" and Rs("qhgj")="1" Then
				mtfz=Rs("mjzf")*Rs("mtbl")/100
				dxfz=Rs("mjzf")*(100-Rs("mtbl"))/100
			End if
			'��Ӳ�󹲼��ķ�ֵ�����ӵ��󹲼�����
			If Rs("gjfs")="3" and Rs("qhgj")="2" Then
				mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100
				dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
			'�������������й������ȫ�ӵ�ģͷ
			If (not (Rs("gjfs")="3")) Then
				mtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + gjfz
				dxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
		Case Else		'09�湲���Ʒ�ģʽ
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
'������ģͷ�Ͷ��ͽṹ����ʱ��ȡ���������һ��
	select case rs("rwlr")
		case "���"
			'����
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ�������',"&Round(mtfz*igysjxs,1)&",'"&now()&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹������',"&Round(dxfz*igysjxs,1)&",'"&now()&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�����������',"&Round(hgjf*igysjxs,1)&",'"&now()&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ�������',"&Round(mtfz*igysjsh,1)&",'"&now()&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹������',"&Round(dxfz*igysjsh,1)&",'"&now()&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�����������',"&Round(hgjf*igysjsh,1)&",'"&now()&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
		case "����"
			'����
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո���',"&Round(mtfz*igyfgxs,1)&",'"&now()&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո���',"&Round(dxfz*igyfgxs,1)&",'"&now()&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո���',"&Round(hgjf*igyfgxs,1)&",'"&now()&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո������',"&Round(mtfz*igyfgsh,1)&",'"&now()&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո������',"&Round(dxfz*igyfgsh,1)&",'"&now()&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո������',"&Round(hgjf*igyfgsh,1)&",'"&now()&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
		case "����"
			'����
			if not(isNull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո���',"&Round(mtfz*igyfcxs,1)&",'"&now()&"',"&ijc&",'"&rs("mtgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո���',"&Round(dxfz*igyfcxs,1)&",'"&now()&"',"&ijc&",'"&rs("dxgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո���',"&Round(hgjf*igyfcxs,1)&",'"&now()&"',"&ijc&",'"&rs("gjgysjr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','ģͷ���ո������',"&Round(mtfz*igyfcsh,1)&",'"&now()&"',"&ijc&",'"&rs("mtgyshr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','���͹��ո������',"&Round(dxfz*igyfcsh,1)&",'"&now()&"',"&ijc&",'"&rs("dxgyshr")&"',"&iGroup&")"
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
				strSql="insert into [mantime] (lsh, rwlr, fz, jssj,  jc, zrr, xz) values ('"&rs("lsh")&"','�������ո������',"&Round(hgjf*igyfcsh,1)&",'"&now()&"',"&ijc&",'"&rs("gjgyshr")&"',"&iGroup&")"
				call xjweb.Exec(strSql,0)
			end if
	end select
	rs.close
end function
%>
