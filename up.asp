<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
CurPage="��������"					'ҳ�������λ��( ��������� �� ���������)yutg
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
     	 Call DeToDB()
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

Function DeToDB()
	'������д�붨���
	Dim mtfz, dxfz, imtjgbl, idxjgbl, isjbl, ishbl, ibombl, hgjf, ifzxs, igyxs
	Dim iGroup, tmpSql, tmpRs, imtrw, idxrw, ijssj, ijhjssj, ikhxs
	hgjf=0 : iGroup=0 : ikhxs=1 : igyxs=1
	'ָ����ơ���ˡ�bom������������ӵ���1
	isjbl=0.72 : ishbl=0.24 : ibombl=0.04

	strSql="select * from [mtask] where datediff('d',sjjssj,'"&now()&"')<50"
	set rs=xjweb.Exec(strSql,1)
	Do while not Rs.eof
		imtrw=Rs("mtrw")
		idxrw=Rs("dxrw")
		ijhjssj=Rs("jhjssj")
		ijssj=Rs("sjjssj")
		if Isnull(ijssj) then ijssj=now()
		ikhxs=1

		ifzxs=Rs("fzxs")
		if imtrw="����" or idxrw="����" Then igyxs=0.33

			if not(isnull(rs("mtgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("mtgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, xz, js, fz, sjfz, jssj, bz) values ('"&rs("lsh")&"','"&rs("mtgysjr")&"',"&iGroup&",'ģͷ�������',"&Round(20*ifzxs*igyxs,1)&","&Round(20*ifzxs*igyxs,1)&",'"&ijssj&"','ʱ��ϵ��:"&ikhxs&"')"
				call xjweb.Exec(strSql,0)
			end if
			if not(isnull(rs("dxgysjr"))) then
				tmpSql="Select [user_group] from [ims_user] where [user_name]='"&rs("dxgysjr")&"'"
				Set tmpRs=xjweb.Exec(tmpSql,1)
				If Not(tmpRs.Eof Or tmpRs.Bof) Then
					iGroup=tmpRs("user_group")
				Else
					iGroup=0
				End If
				tmpRs.Close
				strSql="insert into [reward] (lsh, zrr, xz, js, fz, sjfz, jssj, bz) values ('"&rs("lsh")&"','"&rs("dxgysjr")&"',"&iGroup&",'���͹������',"&Round(30*ifzxs*igyxs,1)&","&Round(30*ifzxs*igyxs,1)&",'"&ijssj&"','ʱ��ϵ��:"&ikhxs&"')"
				call xjweb.Exec(strSql,0)
			end if
	Rs.movenext
	Loop
	rs.close
end function
%>
