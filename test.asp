<%Function DeToDB(lsh)
	'������д�붨���
	Dim mtfz, dxfz, bomfz, ifgshbl, ifcbl, ijc, ijc2, imtjgbl, idxjgbl, ijgshbl, iljshbl, ibombl, iwcsj, hgjf
	Dim iGroup, tmpSql, tmpRs, imtrw, idxrw
	hgjf=0
	'ָ����ơ���ˡ�bom������������ӵ���1
	isjbl=0.72 : ishbl=0.24 : ibombl=0.04

	strSql="select * from [mtask] where lsh='"&lsh&"'"
	set rs=xjweb.Exec(strSql,1)
	if NullToNum(Rs("hgj"))<>0 Then hgjf=200
	if NullToNum(Rs("mtjgbl"))<>0 Then imtjgbl=Rs("mtjgbl")/100
	if NullToNum(Rs("dxjgbl"))<>0 Then idxjgbl=Rs("dxjgbl")/100
	mtfz=Rs("demt")
	dxfz=Rs("dedx")
	imtrw=Rs("mtrw")
	idxrw=Rs("dxrw")
	ijc2=datediff("d",rs("sjjssj"),rs("jhjssj"))
	ijc=0

	Dim ijgsj, isj
	If IsNull(rs("jhjgsj")) Then
		isj=INT(datediff("d", rs("jhkssj"), rs("jhjssj"))/2)
		ijgsj=dateadd("d",isj,rs("jhkssj"))
	else
		ijgsj=rs("jhjgsj")
	End if
'������ģͷ�Ͷ��ͽṹ����ʱ��ȡ���������һ��
			if not(isnull(rs("mtjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("mtjgr")&"',"&iGroup&",'ģͷ�ṹ',"&Round(mtfz*isjbl*imtjgbl,1)&",'"&rs("sjjssj")&")"
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
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("dxjgr")&"',"&iGroup&",'���ͽṹ',"&Round(dxfz*isjbl*idxjgbl,1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjjgr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjjgr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("gjjgr")&"',"&iGroup&",'�󹲼��ṹ',"&Round(hgjf*isjbl*imtjgbl,1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("mtsjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("mtsjr")&"',"&iGroup&",'ģͷ���',"&Round(mtfz*isjbl*(1-imtjgbl),1)&",'"&rs("sjjssj")&")"
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
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("dxsjr")&"',"&iGroup&",'�������',"&Round(dxfz*isjbl*(1-idxjgbl),1)&",'"&rs("sjjssj")&")"
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
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("gjsjr")&"',"&iGroup&",'�󹲼����',"&Round(hgjf*isjbl*(1-imtjgbl),1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("mtjgshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("mtjgshr")&"',"&iGroup&",'ģͷ�ṹ���',"&Round(mtfz*ishbl*imtjgbl,1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("mtsjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("mtsjshr")&"',"&iGroup&",'ģͷ������',"&Round(mtfz*ishbl*(1-imtjgbl),1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxjgshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("dxjgshr")&"',"&iGroup&",'���ͽṹ���',"&Round(dxfz*ishbl*idxjgbl,1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxsjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("dxsjshr")&"',"&iGroup&",'����������',"&Round(dxfz*ishbl*(1-idxjgbl),1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjjgshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjjgshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("gjjgshr")&"',"&iGroup&",'�󹲼��ṹ���',"&Round(hgjf*ishbl*imtjgbl,1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("gjsjshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("gjsjshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("gjsjshr")&"',"&iGroup&",'�󹲼�������',"&Round(hgjf*ishbl*(1-imtjgbl),1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("mtbomr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtbomr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("mtbomr")&"',"&iGroup&",'ģͷBOM',"&Round(mtfz*ibombl,1)&",'"&rs("sjjssj")&")"
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
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("dxbomr")&"',"&iGroup&",'����BOM',"&Round(dxfz*ibombl,1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
			
			'������ˡ�����		
			if not(isnull(rs("mtshr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtshr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				if imtrw="����" Then
					strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("mtshr")&"',"&iGroup&",'ģͷ����',"&Round(mtfz,1)&",'"&rs("sjjssj")&")"				
				else
					strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("mtshr")&"',"&iGroup&",'ģͷ���',"&Round(mtfz*ishbl,1)&",'"&rs("sjjssj")&")"
				End If
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
				if idxrw="����" Then				
					strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("dxshr")&"',"&iGroup&",'���͸���',"&Round(dxfz,1)&",'"&rs("sjjssj")&")"
				else
					strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("dxshr")&"',"&iGroup&",'�������',"&Round(dxfz*ishbl,1)&",'"&rs("sjjssj")&")"
				End if					
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
				strSql="insert into [reward] (lsh, zrr, group, js, fz, sjfz, jssj) values ('"&rs("lsh")&"','"&rs("gjshr")&"',"&iGroup&",'�������',"&Round(hgjf,1)&",'"&rs("sjjssj")&")"
				call xjweb.Exec(strSql,0)
			end if
	rs.close
end function
%>