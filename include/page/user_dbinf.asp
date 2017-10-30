         <%
	'9:05 2007-2-2-星期五
	'本页用到类xj_classs的实例xjweb
	dim c_alluser, c_allzy, c_allzz, c_alljl, c_alltsy, c_alljszd, c_allfs, c_allbc, c_zypx, maxGroup
	dim c_xz0, c_xz1, c_xz2, c_xz3, c_xz4, c_xz5, c_xz6, c_xz7, c_xz8, c_allstat, c_jsb
	dim tmpname, tmpGroup, tmpDept, tmpAble
	c_alluser="": c_allzy="": c_allzz="": c_alljl="" : c_alltsy="" : c_alljszd="" : c_allfs="" : c_allbc=""
	c_xz0="" : c_xz1="" : c_xz2="" : c_xz3="" : c_xz4="" : c_xz5="" : c_xz6="" : c_xz7="" : c_xz8="" : c_allstat=""
	c_jsb="" : c_zypx="" : maxGroup=0
	set rs = xjweb.exec("select user_name, user_able, user_group, user_depart from ims_user where user_name<>'AA' and user_name<>'BB' order by user_depart,user_name",1)
	do while not rs.eof
		tmpname=replace(rs("user_name"),"|||","")
		tmpGroup=Rs("user_group")
		if tmpGroup>maxGroup then maxGroup=tmpGroup
		tmpDept=Rs("user_depart")
		tmpAble=Rs("user_Able")

		'所有人员
		if c_alluser <> "" then
			c_alluser=c_alluser & "|||" & tmpname
		else
			c_alluser=tmpname
		end if
		'经理列表
		if chkuser(3) then
			if c_alljl <> "" then
				c_alljl=c_alljl & "|||" & tmpname
			else
				c_alljl=tmpname
			end if
		end if
		'组长数组
		if chkuser(4) then
			if c_allzz <> "" then
				c_allzz=c_allzz & "|||" & tmpname
			else
				c_allzz=tmpname
			end if
		end if
		'组员数组
		if InStr("145689",ChkJs(tmpAble))>0 then
			if c_allzy <> "" then
				c_allzy=c_allzy & "|||" & tmpname
			else
				c_allzy=tmpname
			end if
			Select Case tmpgroup
				Case 0
					If c_xz0<>"" Then
						c_xz0=c_xz0 & "|||" & tmpname
					Else
						c_xz0=tmpname
				End If
				Case 1
					If c_xz1<>"" Then
						c_xz1=c_xz1 & "|||" & tmpname
					Else
						c_xz1=tmpname
					End If
				Case 2
					If c_xz2<>"" Then
						c_xz2=c_xz2 & "|||" & tmpname
					Else
						c_xz2=tmpname
					End If
				Case 3
					If c_xz3<>"" Then
						c_xz3=c_xz3 & "|||" & tmpname
					Else
						c_xz3=tmpname
					End If
				Case 4
					If c_xz4<>"" Then
						c_xz4=c_xz4 & "|||" & tmpname
					Else
						c_xz4=tmpname
					End If
				Case 5
					If c_xz5<>"" Then
						c_xz5=c_xz5 & "|||" & tmpname
					Else
						c_xz5=tmpname
					End If
				Case 6
					If c_xz6<>"" Then
						c_xz6=c_xz6 & "|||" & tmpname
					Else
						c_xz6=tmpname
					End If
				Case 7
					If c_xz7<>"" Then
						c_xz7=c_xz7 & "|||" & tmpname
					Else
						c_xz7=tmpname
					End If
				Case 8
					If c_xz8<>"" Then
						c_xz8=c_xz8 & "|||" & tmpname
					Else
						c_xz8=tmpname
					End If
			End Select
		end if
		'调试员数组
		if chkuser(6) then
			if instr(tmpname,"调试")=0 Then
				if c_alltsy <> "" then
					c_alltsy=c_alltsy & "|||" & tmpname
				else
					c_alltsy=tmpname
				end if
			end if
		end if

		'复审数组
		if chkuser(8) then
				if c_allfs <> "" then
					c_allfs=c_allfs & "|||" & tmpname
				else
					c_allfs=tmpname
				end if
		end if

		'编程数组
		if chkuser(9) then
				if c_allbc <> "" then
					c_allbc=c_allbc & "|||" & tmpname
				else
					c_allbc=tmpname
				end if
		end if

		'所有统计人员,包括TT、TB调试员
		if InStr("12345689",ChkJs(tmpAble))>0 Or chkuser(10)  then
			if c_allstat <> "" then
				c_allstat=c_allstat & "|||" & tmpname
			else
				c_allstat=tmpname
			end if
		end if

		'所有技术部成员,不包括TT、TB调试员
		if tmpDept="技术部" and instr(tmpname,"调试员")=0 then
			if c_jsb <> "" then
				c_jsb=c_jsb & "|||" & tmpname
			else
				c_jsb=tmpname
			end if
		end if

		rs.movenext
	loop
	rs.close

	set rs = xjweb.exec("select user_name, user_able, user_group, user_depart from ims_user where user_name<>'AA' and user_name<>'BB' and user_depart='技术部' order by user_group",1)
	do while not rs.eof
		tmpname=replace(rs("user_name"),"|||","")
		tmpGroup=Rs("user_group")
		tmpDept=Rs("user_depart")
		tmpAble=Rs("user_Able")
		'所有统计人员,包括TT、TB调试员按组排序
		if InStr("1456",ChkJs(tmpAble))>0 then
			if c_zypx <> "" then
				c_zypx=c_zypx & "|||" & tmpname
			else
				c_zypx=tmpname
			end if
		end if
		rs.movenext
	loop
	rs.close

	If c_alluser="" Then c_alluser=" "
	If c_allzy="" Then c_allzy=" "
	If c_allzz="" Then c_allzz=" "
	If c_alljl="" Then c_alljl=" "
	If c_alltsy="" Then c_alltsy=" "
	If c_allfs="" Then c_allfs=" "
	If c_allbc="" Then c_allbc=" "
	If c_xz0="" Then c_xz0=" "
	If c_xz1="" Then c_xz1=" "
	If c_xz2="" Then c_xz2=" "
	If c_xz3="" Then c_xz3=" "
	If c_xz4="" Then c_xz4=" "
	If c_xz5="" Then c_xz5=" "
	If c_xz6="" Then c_xz6=" "
	If c_xz7="" Then c_xz7=" "
	If c_xz8="" Then c_xz8=" "
	If c_allstat="" Then c_allstat=" "
	If c_zypx="" Then c_zypx=" "
	If c_jsb="" Then c_jsb=" "

	c_alluser = split(c_alluser, "|||")
	c_allzy = split(c_allzy, "|||")
	c_allzz = split(c_allzz, "|||")
	c_alljl = split(c_alljl, "|||")
	c_alltsy = split(c_alltsy, "|||")
	c_allfs = split(c_allfs, "|||")
	c_allbc = split(c_allbc, "|||")
	c_xz0=split(c_xz0,"|||")
	c_xz1=split(c_xz1,"|||")
	c_xz2=split(c_xz2,"|||")
	c_xz3=split(c_xz3,"|||")
	c_xz4=split(c_xz4,"|||")
	c_xz5=split(c_xz5,"|||")
	c_xz6=split(c_xz6,"|||")
	c_xz7=split(c_xz7,"|||")
	c_xz8=split(c_xz8,"|||")
	c_allstat=split(c_allstat,"|||")
	c_zypx=split(c_zypx,"|||")
	c_jsb=split(c_jsb,"|||")

function chkuser(num)
	'num 为权限位的位数,如第四位则str=4 , rs: 数据库记录集
	chkuser=false
	if not isnumeric(num) then num=0
	num=cint(num)
	if num<0 then exit function
	if num>len(rs("user_able")) then exit function
	if mid(rs("user_able"),num,1)>0 then chkuser=true
end function

Function ChkJs(str)
	'str 为权限000001000000000
	ChkJs=0
	If Len(str)<15 Then Exit Function
	dim i
	For i=1 To Len(str)
		If Mid(str,i,1)=1 Then ChkJs=i : Exit For	'只取每人的最高角色,如你同时是组长和组员,则只取组长
	Next
End Function
%>