<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="��ֵͳ�� �� �鿴Ա��ϵ��"					'ҳ�������λ��( ��ֵͳ�� �� �鿴Ա��ϵ��)
strPage="mtstat"
xjweb.header()
Call TopTable()

Dim iyear, imonth, dtstart, dtend, irwzf, iaddfz, zcount, icount, ilxrwzf, zrwwcl, zgroup
Dim zuser, zrwfz, zrwxs, zzlxs, zgkxs, zbmxs, zjbgz, zjxgz,zyfgz, zbeiz, ygxsRs, m, zbasicwg
zjbgz=0
zjxgz=0
zyfgz=0
zgroup=0
zbasicwg=0
zcount=1
icount=1
iyear = request("searchy")
imonth = request("searchm")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)
dtend=cdate(iyear&"��"&imonth&"��1��")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
dtstart=cdate(iyear&"��"&imonth&"��1��")

'���忼���õı���
	Dim kpf(30), kpif(10), ics(10), kpzf, kpxr
	kpxr=Array("")

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchMantime()%></td>
  </tr>
 </Table>
<%Call YgxsDisplay()
      Response.Write(XjLine(10,"100%",""))
End Sub

Function SearchMantime()
%>
<table cellpadding=2 cellspacing=0>
  <form action=<%=request.servervariables("script_name")%> method=get>
    <tr>
      <td> ��ѡ��:
        <select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&bybmxs="+this.form.bybmxs.value);'>
          <%for i = year(now) - 3 to year(now)%>
          <option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        ��
        <select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&bybmxs="+this.form.bybmxs.value);'>
          <%for i = 1 to 12%>
          <option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        ��&nbsp;&nbsp;
        <label>���²���ϵ����
          <input type="text" name="bybmxs" size="4"  onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&bybmxs="+this.form.bybmxs.value);'>
          &nbsp;&nbsp; </label>
        <label>��������ʣ�
          <input type="text" name="rwwcl" size="4">
        </label>
        &nbsp;&nbsp;
        <input type="submit" value=" ȷ �� "></td>
    </tr>
  </form>
</table>
<%
End Function

Function YgxsDisplay()		'��ʾ�б�
		Call TbTopic("������������Ա" & iyear & "��" & imonth & "�¿��˻��ܱ�")
		If Request("rwwcl")="" Then zrwwcl=1.0 Else zrwwcl=Request("rwwcl") End if
		%>
<table cellpadding=2 cellspacing=0 class="xtable" width="<%=web_info(8)%>">
<THEAD>
  <tr>
    <th class=th width="8%">ID</th>
    <th class=th width="8%">��Ա����</th>
    <th class=th width="8%">�����ֵ</th>
    <th class=th width="8%">����ָ��</th>
    <th class=th width="8%">����</th>
    <th class=th width="8%">����</th>
    <th class=th width="8%">�ۺ�</th>
    <th class=th width="8%">����ϵ��</th>
    <th class=th width="10%">��������</th>
    <th class=th width="10%">��Ч����</th>
    <th class=th width="10%">Ӧ������</th>
  </tr>
  <tr>
  	<td colspan="11" class=rtd>���²������������=<%=zrwwcl%></td>
  </tr>
  </THEAD>
  <%
		Dim strColor
		strColor=-1
		If Request("bybmxs")="" Then zbmxs=1.0 Else zbmxs=Request("bybmxs") End if
		strSql="select * from [ims_user] where user_depart='������' and user_group<>0 and user_able<>'010000000000000' and Instr('AABBTB����Ա',user_name)=0 order by user_group,user_able"
		Set ygxsRs=xjweb.Exec(strSql, 1)
		Do While Not ygxsRs.eof	or ygxsRs.Bof
		zuser=ygxsRs("user_name")
		If zgroup<>ygxsRs("user_group") Then
			strColor=-1*strColor
			zgroup=ygxsRs("user_group")
		End If
		zgroup=ygxsRs("user_group")
		zbasicwg=ygxsRs("user_basicwage")		
		zrwfz=0 : zrwxs=0 : zzlxs=0 : zgkxs=0 : zjbgz=0 : zyfgz=0 : zbeiz="" : irwzf=0 : ilxrwzf=0 : iaddfz=0
		kpzf=0
		for i=0 to 29
			kpf(i)=0
		next
		for i=0 to 9
			kpif(i)=0
		next
		for i=0 to 9
			ics(i)=0
		next
		Call YgxsStat()
		zrwxs=FormatNumber(zrwxs,2)
	%>
<TBODY>
  <tr <%If strColor=1 Then%>bgcolor="#D6D7EF"<%End If%>>
    <td class=ctd width="8%"><%=zcount%></td>
    <td class=ctd width="8%"><%=zuser%></td>
    <td class=ctd width="8%"><%=zrwfz%>&nbsp;</td>
    <td class=ctd width="8%"><%=zbasicwg%>&nbsp;</td>
    <td class=ctd width="8%"><%=zrwxs%>&nbsp;</td>
    <td class=ctd width="8%"><%=zzlxs%>&nbsp;</td>
    <td class=ctd width="8%"><%=zgkxs%>&nbsp;</td>
    <td class=ctd width="8%"><%=zbmxs%></td>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="10%">&nbsp;</td>
    <td class=ctd width="10%">&nbsp;</td>
  </tr>
</TBODY>
  <%
		zcount = zcount + 1
		ygxsRs.movenext
		loop
		ygxsRs.close
%>
<TFOOT>
<TR>
<TD class=rtd colspan=11>
The End.
</TD>
</TR>
</TFOOT>
</table>
<%
End Function

Function YgxsStat()
	strSql="Select * from [ims_user] where [user_name]='"&zuser&"'"
	Set Rs=xjweb.Exec(strSql,1)
	Dim tmpCount, tmpGroup, tmpAble, ilxrwzf
	tmpCount=1
	tmpGroup=Rs("user_Group")
	tmpAble=Rs("user_Able")
	Rs.Close

	If InStr("4568",ChkJs(tmpAble))>0 Then		'�ж��ǲ�����Ա�����Ա
		'1--�����ֵ
		strSql="select * from [mantime] where zrr='"&zuser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			irwzf=irwzf+Round(Rs("fz"),1)
			Rs.movenext
		Loop
		Rs.close
		'2---���������ֵ
		strSql="select * from [ftask] where zrr='"&zuser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			ilxrwzf=ilxrwzf+Rs("zf")
			Rs.movenext
		Loop
		Rs.close
		'3---ͳ���ܷ�,
			If Fix(ilxrwzf + irwzf)<(ilxrwzf + irwzf) Then
				zrwfz=Fix(ilxrwzf + irwzf) + 1
			Else
				zrwfz=Fix(ilxrwzf + irwzf)
			End If
	End If
	icount=1
	Select Case ChkJs(tmpAble)
		Case 6	'����Ա
					kpif(0)=statkpfz("�ӳ�", 0)
					kpif(1)=statkpfz("��ǰ", 0)
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					kpif(0)=statkpfz("5����������δ����", 0)
					kpif(1)=statkpfz("�������·�����ʱ", 0)
					kpif(2)=statkpfz("�����鱨¼�벻��ʱ", 0)
					kpif(3)=statkpfz("����ͼֽǩ�𡢸��Ĳ�����", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(2)<-8 Then kpf(2)=-8
					kpif(0)=statkpfz("����ͼδ��ʱ�浵", 0)
					kpif(1)=statkpfz("������ԭ���������", 0)
					kpif(1)=kpif(1)+statkpfz("���ԭ���������", 0)
					kpif(2)=statkpfz("������ԭ���������", 0)
					kpif(2)=kpif(2)+statkpfz("���ԭ���������", 0)
					kpf(3)=kpif(0)+kpif(1)+kpif(2)
					If kpf(3)<-20 Then kpf(3)=-20
					kpif(0)=statkpfz("���Գ����������", 0)
					kpif(1)=statkpfz("�������ڶ��С����", 0)
					kpif(2)=statkpfz("ģ�ߵ���δ�ϸ���", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-6 Then kpf(4)=-6
					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("ֵ�����,����", 0)
					kpif(2)=statkpfz("���治��,�°����δ�ء���δ��", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
					kpif(0)=statkpfz("�����ӷ���", 0)
					kpif(1)=statkpfz("��������", 0)
					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>zbasicwg Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(zrwfz/zbasicwg * 50),1)
				End If
				If zrwfz<zbasicwg Then
					zrwxs=round((zrwfz/zbasicwg * 50)/100,2)
				Else
					zrwxs=round(((50+((zrwfz-zbasicwg)/zbasicwg*50*1.25))/100),2)
				End If
				zzlxs=round(kpzf/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
				zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 8		'������Ա
					kpif(0)=statkpfz("��Ʊ�׼���淶ά������ʱ", 0)
					kpif(1)=statkpfz("���Է������⴦����ʱ", 0)
					kpif(2)=statkpfz("Ӫ������֧�ֲ���ʱ", 0)
					kpif(3)=statkpfz("������������", 0)
					kpf(0)=50 + kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(0)<0 Then kpf(0)=0

					kpif(0)=statkpfz("����ԭ�����ⲿͶ��", 0)
					kpif(1)=statkpfz("��������δ���涨ִ��", 0)
					kpif(2)=statkpfz("���ԭ��������ʧ��ǧԪ", 0)
					kpf(1)=30 + kpif(0)+kpif(1)+kpif(2)
					If kpf(1)<0 Then kpf(1)=0

					kpif(0)=statkpfz("����Ľ�����ȡ�ó�Ч", 0)
					kpif(1)=statkpfz("�����е���������", 0)
					kpif(2)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(3)=statkpfz("�����ӷ���", 0)
					kpif(4)=statkpfz("5S�������", 0)
					kpf(2)=20 + kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4)
					If kpf(2)<0 Then kpf(2)=0					
					
				for i=0 to 1
					kpzf=kpzf+kpf(i)
				next
				zrwxs=round(kpzf/100,2)
				zzlxs=round(kpf(2)/100,2)
				zgkxs=round(zrwxs+zzlxs,2)

		Case Else	'��Ա
					If zrwfz<zbasicwg Then
						kpf(0)=round((zrwfz/zbasicwg * 35),1)
					Else
						kpf(0)=round((35+((zrwfz-zbasicwg)/zbasicwg*35*1.25)),1)
					End If
					kpif(0)=statkpfz("������������", 0)
					kpf(1)=kpif(0) + 5
					If kpf(1)<0 Then kpf(1)=0
					
					kpif(0)=statkpfz("�ӳ�", 0)
					kpf(2)=kpif(0) + 10
					If kpf(2)<0 Then kpf(2)=0
					
					kpif(0)=statkpfz("���ԭ���������", 0)
					kpif(1)=statkpfz("���ԭ���������", 0)
					kpif(2)=statkpfz("���ԭ���������", 0)
					kpif(3)=statkpfz("���ڵ������ڶ����", 0)
					kpif(4)=statkpfz("���ڵ��Զ��ڶ����", 0)
					kpif(5)=statkpfz("��������δ���涨ִ��", 0)
					kpif(6)=statkpfz("���ԭ��������ʧ��ǧԪ", 0)
					kpf(3)=30+kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)+kpif(5)+kpif(6)
					if kpf(3)<0 Then kpf(3)=0
					
					kpif(0)=statkpfz("����Ľ�����ȡ�ó�Ч", 0)
					kpif(1)=statkpfz("�����е���������", 0)
					kpif(2)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(3)=statkpfz("�����ӷ���", 0)
					kpif(4)=statkpfz("5S�������", 0)
					kpf(4)=20+kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4)
					If kpf(4)<0 Then kpf(4)=0
					
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				if kpzf>50 Then kpzf=50
				kpzf=kpzf+kpf(3)
				zrwxs=round(kpzf/100,2)
				zzlxs=round(kpf(4)/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
	End Select
	Erase kpf
End Function

Function ChkJs(str)
	'str ΪȨ��000001000000000
	ChkJs=0
	If Len(str)<15 Then Exit Function
	dim i
	if Mid(str,8,1)=1 Then	'����������Ա���ȼ�
		ChkJs=8
	Else
		For i=1 To Len(str)
			If Mid(str,i,1)=1 Then ChkJs=i : Exit For	'ֻȡÿ�˵���߽�ɫ,����ͬʱ���鳤����Ա,��ֻȡ�鳤
		Next
	End If
End Function

Function statkpjfz(kp_item,zrrjs,i)
	Dim PjCs,tmpRs
	statkpjfz=0
	strSql="select * from [kp_jsb] where [kp_item] like '%"&kp_item&"%' and [kp_zrrjs] like '%"&zrrjs&"%' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	Set tmpRs=xjweb.Exec(strSql, 1)
	do while not tmpRs.eof
		statkpjfz=statkpjfz + tmpRs("kp_uprice") * tmpRs("kp_mul")
		tmpRs.movenext
	loop
	tmpRs.close
	strSql="Select * from [ims_user] where mid(user_able,4,1)>0 and user_Group>0 and user_Group<>5"
	Call xjweb.Exec("",-1)
	Set tmpRs=Server.CreateObject("ADODB.RECORDSET")
	tmpRs.open strSql,Conn,1,3
		PjCs=tmpRs.RecordCount
	tmpRs.close
	statkpjfz=Round(statkpjfz/PjCs,2)
End Function

Function statkpfz(kp_item, i)
	statkpfz=0
	Dim tmpRs
	Select Case i
		Case 0		'����Ա����ͳ��
			strSql="select * from [kp_jsb] where [kp_zrr]='"&zuser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case -1		'�����ν���ͳ��
			strSql="select * from [kp_jsb] where [kp_item]='"&kp_item&"'  and [kp_kpr]<>" & zuser & " and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'���鳤����ͳ��
			strSql="select * from [kp_jsb] where [kp_group]="&i&"  and [kp_kpr]<>" & zuser & " and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End Select

	Set tmpRs=xjweb.Exec(strSql, 1)

	do while not tmpRs.eof
		statkpfz=statkpfz + tmpRs("kp_uprice") * tmpRs("kp_mul")
		tmpRs.movenext
	loop
	statkpfz=round(statkpfz,2)
	tmpRs.close
	set tmprs=nothing
End Function

Function diskpItem(arg1,arg2,arg3,arg4)
	icount=icount+1
	dim tmpcs, tmpkpf
	tmpcs=statkpcs(arg1, "", arg4)
	tmpkpf=tmpcs*arg3*-1
	kpf(icount-1)=tmpkpf
	If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
End Function

Function diskpItemM(arg1,arg2,arg3,arg4,arg5)
	icount=icount+1
	dim tmpcs, tmpkpf, temparg
	temparg=arg1
	tmpcs=0
	tmpkpf=0
	If Instr(arg1,"ԭ���������")>0 Then temparg="ԭ���������"
	If Instr(arg1,"ԭ���������")>0 Then temparg="ԭ���������"
	strSql=""
	If arg4="" Then
		strSql="select * from [kp_jsb] where kp_item like '%"&temparg&"%' and kp_zrrjs='"&arg5&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	else
		strSql="select * from [kp_jsb] where kp_group="&arg4&" and kp_item like '%"&temparg&"%' and kp_zrrjs='"&arg5&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End If
	Set Rs = Server.CreateObject("ADODB.RECORDSET")
	Call xjweb.Exec("",-1)
	Rs.open strSql,Conn,3,3
  	tmpcs=rs.recordcount
	tmpkpf=tmpcs*arg3*-1
	kpf(icount-1)=tmpkpf
	If arg2<>"" Then
		If kpf(icount-1)<arg2*-1 Then
			 kpf(icount-1)=arg2*-1
		End If
	End If
	'�鳤���ޡ�����������
	Rs.Close
End Function
%>
