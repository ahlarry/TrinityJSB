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
Dim zuser, zrwfz, zrwxs, zzlxs, zgkxs, zbmxs, zjbgz, zjxgz,zyfgz, zbeiz, ygxsRs ,m
zjbgz=0
zjxgz=0
zyfgz=0
zgroup=0
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
    <th class=th width="10%">��Ա����</th>
    <th class=th width="10%">�����ֵ</th>
    <th class=th width="10%">��������</th>
    <th class=th width="10%">���Կ���</th>
    <th class=th width="10%">�ۺϿ���</th>
    <th class=th width="10%">����ϵ��</th>
    <th class=th width="10%">��������</th>
    <th class=th width="10%">��Ч����</th>
    <th class=th width="10%">Ӧ������</th>
  </tr>
  <tr>
  	<td colspan="10" class=rtd>���²������������=<%=zrwwcl%></td>
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
    <td class=ctd width="10%"><%=zuser%></td>
    <td class=ctd width="10%"><%=zrwfz%>&nbsp;</td>
    <td class=ctd width="10%"><%=zrwxs%>&nbsp;</td>
    <td class=ctd width="10%"><%=zzlxs%>&nbsp;</td>
    <td class=ctd width="10%"><%=zgkxs%>&nbsp;</td>
    <td class=ctd width="10%"><%=zbmxs%></td>
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
<TD class=rtd colspan=10>
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

	If InStr("5689",ChkJs(tmpAble))>0 Then		'�ж��ǲ�����Ա�����Ա
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
	else
		If InStr("4",ChkJs(tmpAble))>0 Then		'�ж��ǲ����鳤
			'1--ͳ��С���Ա����
			strSql="Select * from [ims_user] where [user_group]="&tmpGroup
			Call xjweb.Exec("",-1)
			Set Rs=Server.CreateObject("ADODB.RECORDSET")
			Rs.open strSql,Conn,1,3
			tmpCount=Rs.RecordCount
			Rs.Close
			'2--�����ֵ
			strSql="select * from [mantime] where [xz]="&tmpGroup&" and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
			Set Rs=xjweb.Exec(strSql, 1)
			Do While Not Rs.eof
				irwzf=irwzf+Round(Rs("fz"),1)
				Rs.movenext
			Loop
			Rs.close
			'3---���������ֵ
			strSql="select * from [ftask] where [xz]="&tmpGroup&" and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
			Set Rs=xjweb.Exec(strSql, 1)
			Do While Not Rs.eof
				ilxrwzf=ilxrwzf+Rs("zf")
				Rs.movenext
			Loop
			Rs.close
			'4---ͳ���ܷ�
			zrwfz=Round((ilxrwzf + irwzf)/tmpCount,1)
		End If
	End If
	zjxgz=1800
	icount=1
	Select Case ChkJs(tmpAble)
		Case 3	'����
			%>
<%Call diskpItem("��ƽ����ļ�δ����",5.0,2.5, 0)%>
<%Call diskpItem("���Խ����ļ�δ����",5.0,2.5, 0)%>
<%Call diskpItem("�����ļ�ǩ�𡢸��Ĳ�����",2.0,1.0, -1)%>
<%Call diskpItem("�ļ�����δ�浵",2.0,2.0, 0)%>
<%Call diskpItem("��Ŀ����������",6.0,2.0, -1)%>
<%Call diskpItem("��Ʒ�����¡���δ��ʱ�ϱ�",2.0,1.0, -1)%>
<%Call diskpItem("��Ʒ����ʵʩ�ƻ�ά�������",2.0,1.0, -1)%>
<%Call diskpItem("��ƹ淶ά��",4.0,2.0, 0)%>
<%Call diskpItem("��ͬ���������Ͳ����������ظ�",6.0,2.0, -1)%>
<%Call diskpItem("���ڵ�����Ϣ����",4.0,1.0, -1)%>
<%Call diskpItem("ģ�ߵ���δ�ϸ���",4.0,2.0, 0)%>
<%Call diskpItem("�ͻ���������Ͷ���뱧Թ",4.0,2.0, 0)%>
<%Call diskpItemM("���ԭ���������",8.0,0.4, "", "���")%>
<%Call diskpItemM("���ԭ���������",6.0,0.4, "", "���")%>
<%Call diskpItem("��׼�ṹ����δ�������",4.0,2.0, -1)%>
<%Call diskpItem("���ͷ���ǻ���ṹ���ٶȺ������鲻��",6.0,3.0, -1)%>
<%Call diskpItem("�ӿڼ����Ȱ彺��Ⱥ������鲻��",4.0,2.0, -1)%>
<%Call diskpItem("ͬ���β�Ʒ��ͬ��λģ����Ʋ�һ��",4.0,2.0, -1)%>
<%Call diskpItem("���׼ʱ�����",12.0,2.0, 0)%>
<%Call diskpItem("�������ⲻ��ʱ",1.0,1.0, -1)%>
<%Call diskpItem("����Э��",2.0,2.0, 0)%>
<%Call diskpItem("�����ձ��ܱ��±�",1.0,1.0, 0)%>
<%Call diskpItem("�·�����������Ŀ",1.0,1.0, 0)%>
<%Call diskpItem("���������ƻ����",1.0,1.0, 0)%>
<%Call diskpItem("����Ԥ����ʩ��ִ��",1.0,1.0, 0)%>
<%Call diskpItem("ʹ�÷���Ч�ļ�",1.0,1.0, 0)%>
<%Call diskpItem("�ֳ���������˾ͨ��",1.0,1.0, 0)%>
<%Call diskpItem("Ա����ѵ����������",1.0,1.0, 0)%>
<%	for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				zjxgz=2800
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=zzlxs
			'	zyfgz=zjxgz*zgkxs*zrwwcl*zbmxs+zjbgz

		Case 4	'�鳤
					If zrwfz<300 Then
						kpf(0)=round((zrwfz/300 * 50),1)
					Else
						kpf(0)=round((50+((zrwfz-300)/300*50*1.25)),1)
					End If
					ics(0)=statkpcs("��ǰ", "", 0)
					ics(1)=statkpcs("�ӳ�", "", 0)
					kpif(0)=statkpfz("��ǰ", 0)
					kpif(1)=statkpfz("�ӳ�", 0)
					kpf(1)=kpif(0) + kpif(1)
					if kpf(1)<-10 Then kpf(1)=-10
					if kpf(1)>10 Then kpf(1)=10
					ics(0)=statkpcs("ģ����ƺ������鲻��", "", 0)
					ics(1)=statkpcs("���ԭ���������", "���", 0)+statkpcs("���ԭ���������", "���", tmpGroup)
					ics(2)=statkpcs("���ԭ���������", "���", 0)+statkpcs("���ԭ���������", "���", tmpGroup)
					ics(3)=statkpcs("���ԭ���������", "���", 0)+statkpcs("���ԭ���������", "���", tmpGroup)
					ics(4)=statkpcs("���ڵ������ڶ����", "", 0)
					ics(5)=statkpcs("���ڵ��Զ��ڶ����", "", 0)
					ics(6)=statkpcs("��ͬ���������Ͳ����������ظ�����", "", 0)
					kpif(0)=statkpfz("ģ����ƺ������鲻��", 0)
					kpif(1)=statkpfz("���ԭ���������", 0) - ics(1)
					kpif(2)=statkpfz("���ԭ���������", 0) - ics(2)*2
					kpif(3)=statkpfz("���ԭ���������", 0) - ics(3)*4
					kpif(4)=statkpfz("���ڵ������ڶ����", 0)
					kpif(5)=statkpfz("���ڵ��Զ��ڶ����", 0)
					kpif(6)=statkpfz("��ͬ���������Ͳ����������ظ�����", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4) + kpif(5) + kpif(6)
					if kpf(2)<-20 Then kpf(2)=-20
					if kpf(2)>20 Then kpf(2)=20
					ics(0)=statkpcs("��ƹ淶ά����ʱ��", "", 0)
					ics(1)=statkpcs("��Ŀʵʩ�ƻ�δ��ʱ���", "", 0)
					ics(2)=statkpcs("���ڼ�������ά������ʱ", "", 0)
					ics(3)=statkpcs("��׼�ṹ����δ�������", "", 0)
					ics(4)=statkpcs("�ֳ������������", "", 0)
					ics(5)=statkpcs("�����ӷ���", "", 0)
					kpif(0)=statkpfz("��ƹ淶ά����ʱ��", 0)
					kpif(1)=statkpfz("��Ŀʵʩ�ƻ�δ��ʱ���", 0)
					kpif(2)=statkpfz("���ڼ�������ά������ʱ", 0)
					kpif(3)=statkpfz("��׼�ṹ����δ�������", 0)
					kpif(4)=statkpfz("�ֳ������������", 0)
					kpif(5)=statkpfz("�����ӷ���", 0)
					kpf(3)=kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4) + kpif(5)
					if kpf(3)<-20 Then kpf(3)=-20
					if kpf(3)>20 Then kpf(3)=20
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>300 Then
					kpzf=kpzf+30
				else
					kpzf=kpzf+(zrwfz/300 * 30)
				End If
				zrwxs=round(kpzf/100,2)
				zzlxs=round((20+kpf(3))/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
				zjxgz=2000
		Case 5	'��Ա
					If zrwfz<300 Then
						kpf(0)=round((zrwfz/300 * 50),1)
					Else
						kpf(0)=round((50+((zrwfz-300)/300*50*1.25)),1)
					End If
					ics(0)=statkpcs("�ӳ�", "", 0)
					ics(1)=statkpcs("��ǰ", "", 0)
					kpif(0)=statkpfz("�ӳ�", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("��ǰ", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("���ԭ���������", "", 0)
					ics(1)=statkpcs("���ԭ���������", "", 0)
					ics(2)=statkpcs("���ԭ���������", "", 0)
					ics(3)=statkpcs("���ڵ������ڶ����", "", 0)
					ics(4)=statkpcs("���ڵ��Զ��ڶ����", "", 0)
					kpif(0)=statkpfz("���ԭ���������", 0)
					kpif(1)=statkpfz("���ԭ���������", 0)
					kpif(2)=statkpfz("���ԭ���������", 0)
					kpif(3)=statkpfz("���ڵ������ڶ����", 0)
					kpif(4)=statkpfz("���ڵ��Զ��ڶ����", 0)
					kpf(2)=kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)
					If kpf(2)<-20 Then kpf(2)=-20
					If kpf(2)>20 Then kpf(2)=20
					ics(0)=statkpcs("����Ľ�����ȡ�ó�Ч", "", 0)
					ics(1)=statkpcs("����������鲢������", "", 0)
					kpif(0)=statkpfz("����Ľ�����ȡ�ó�Ч", 0)
					kpif(1)=statkpfz("����������鲢������", 0)
					kpf(3)=kpif(0) + kpif(1)
					ics(0)=statkpcs("�ϰ����빤���޹�", "", 0)
					ics(1)=statkpcs("�����ӷ���", "", 0)
					ics(2)=statkpcs("�����е���������", "", 0)
					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("�����ӷ���", 0)
					kpif(2)=statkpfz("�����е���������", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-20 Then kpf(4)=-20
					If kpf(4)>20 Then kpf(4)=20
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>300 Then
					kpzf=kpzf+30
				else
					kpzf=kpzf+(zrwfz/300 * 30)
				End If
				zrwxs=round(kpzf/100,2)
				zzlxs=round((20+kpf(3)+kpf(4))/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
		Case 6	'����Ա
					ics(0)=statkpcs("�ӳ�", "", 0)
					ics(1)=statkpcs("��ǰ", "", 0)
					kpif(0)=statkpfz("�ӳ�", 0)
					kpif(1)=statkpfz("��ǰ", 0)
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("5����������δ����", "", 0)
					ics(1)=statkpcs("�������·�����ʱ", "", 0)
					ics(2)=statkpcs("�����鱨¼�벻��ʱ", "", 0)
					ics(3)=statkpcs("����ͼֽǩ�𡢸��Ĳ�����", "", 0)
					kpif(0)=statkpfz("5����������δ����", 0)
					kpif(1)=statkpfz("�������·�����ʱ", 0)
					kpif(2)=statkpfz("�����鱨¼�벻��ʱ", 0)
					kpif(3)=statkpfz("����ͼֽǩ�𡢸��Ĳ�����", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(2)<-8 Then kpf(2)=-8
					ics(0)=statkpcs("����ͼδ��ʱ�浵", "", 0)
					ics(1)=statkpcs("������ԭ���������", "", 0)
					ics(1)=ics(1)+statkpcs("���ԭ���������", "", 0)
					ics(2)=statkpcs("������ԭ���������", "", 0)
					ics(2)=ics(2)+statkpcs("���ԭ���������", "", 0)
					kpif(0)=statkpfz("����ͼδ��ʱ�浵", 0)
					kpif(1)=statkpfz("������ԭ���������", 0)
					kpif(1)=kpif(1)+statkpfz("���ԭ���������", 0)
					kpif(2)=statkpfz("������ԭ���������", 0)
					kpif(2)=kpif(2)+statkpfz("���ԭ���������", 0)
					kpf(3)=kpif(0)+kpif(1)+kpif(2)
					If kpf(3)<-20 Then kpf(3)=-20
					ics(0)=statkpcs("���Գ����������", "", 0)
					ics(1)=statkpcs("�������ڶ��С����", "", 0)
					ics(2)=statkpcs("ģ�ߵ���δ�ϸ���", "", 0)
					kpif(0)=statkpfz("���Գ����������", 0)
					kpif(1)=statkpfz("�������ڶ��С����", 0)
					kpif(2)=statkpfz("ģ�ߵ���δ�ϸ���", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-6 Then kpf(4)=-6
					ics(0)=statkpcs("�ϰ����빤���޹�", "", 0)
					ics(1)=statkpcs("ֵ�����,����", "", 0)
					ics(2)=statkpcs("���治��,�°����δ�ء���δ��", "", 0)
					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("ֵ�����,����", 0)
					kpif(2)=statkpfz("���治��,�°����δ�ء���δ��", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
					ics(0)=statkpcs("�����ӷ���", "", 0)
					ics(1)=statkpcs("��������", "", 0)
					kpif(0)=statkpfz("�����ӷ���", 0)
					kpif(1)=statkpfz("��������", 0)
					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>300 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(zrwfz/300 * 50),1)
				End If
				If zrwfz<300 Then
					zrwxs=round((zrwfz/300 * 50)/100,2)
				Else
					zrwxs=round(((50+((zrwfz-300)/300*50*1.25))/100),2)
				End If
				zzlxs=round(kpzf/100,2)
				zgkxs=round(zrwxs+zzlxs,2)
				zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 7			'ͼ������Ա
			zjxgz=600
			zzlxs=1.00
			zgkxs=1.00
		'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz


		Case 8		'����Ա
					ics(0)=statkpcs("�ӳ�", "", 0)
					ics(1)=statkpcs("��ǰ", "", 0)
					kpif(0)=statkpfz("�ӳ�", 0)
					kpif(1)=statkpfz("��ǰ", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("��Ʒ�����¼ƻ�����δ���", "", 0)
					ics(1)=statkpcs("��Ʒ�����¡���δ��ʱ�ϱ�", "", 0)
					ics(2)=statkpcs("��Ʒ����ʵʩ�ƻ�ά�������", "", 0)
					kpif(0)=statkpfz("��Ʒ�����¼ƻ�����δ���", 0)
					kpif(1)=statkpfz("��Ʒ�����¡���δ��ʱ�ϱ�", 0)
					kpif(2)=statkpfz("��Ʒ����ʵʩ�ƻ�ά�������", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2)
					If kpf(2)<-5 Then kpf(2)=-5
					ics(0)=statkpcs("�����ļ����Ĳ�����δ�����������", "", 0)
					kpif(0)=statkpfz("�����ļ����Ĳ�����δ�����������", 0)
					ics(1)=statkpcs("�����Ƶ���ҵָ���龭��֤��δ��ʱ�޸�", "", 0)
					kpif(1)=statkpfz("�����Ƶ���ҵָ���龭��֤��δ��ʱ�޸�", 0)
					ics(2)=statkpcs("�����ļ�ǩ������", "", 0)
					kpif(2)=statkpfz("�����ļ�ǩ������", 0)
					ics(3)=statkpcs("��ǩ�����ڲ�ʵ��������������", "", 0)
					kpif(3)=statkpfz("��ǩ�����ڲ�ʵ��������������", 0)
					ics(4)=statkpcs("���մ�����ɷ��ޣ�©�����գ�", "", 0)
					kpif(4)=statkpfz("���մ�����ɷ��ޣ�©�����գ�", 0)
					ics(5)=statkpcs("���մ�����ɱ���", "", 0)
					kpif(5)=statkpfz("���մ�����ɱ���", 0)
					ics(6)=statkpcs("δ���淶ִ��", "", 0)
					kpif(6)=statkpfz("δ���淶ִ��", 0)
					ics(7)=statkpcs("������ɳ��䶩�������´�", "", 0)
					kpif(7)=statkpfz("������ɳ��䶩�������´�", 0)
					ics(8)=statkpcs("��ͬ�����ظ�����", "", 0)
					kpif(8)=statkpfz("��ͬ�����ظ�����", 0)
					ics(9)=statkpcs("��ƽṹ�����ϡ��ȴ�������Դ���δ��ʱ��ӳ", "", 0)
					kpif(9)=statkpfz("��ƽṹ�����ϡ��ȴ�������Դ���δ��ʱ��ӳ", 0)
					kpf(3)=kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)+kpif(5)+kpif(6)+kpif(7)+kpif(8)+kpif(9)
					If kpf(3)<-20 Then kpf(3)=-20
					ics(0)=statkpcs("��������ǹ���������ɵļӹ�����", "", 0)
					ics(1)=statkpcs("��������ӹ����ոĽ�������������", "", 0)
					ics(2)=statkpcs("����ר�üо߳ɹ����ڼƻ���������", "", 0)
					ics(3)=statkpcs("�½�������Ա���ܿ��˲��ϸ�", "", 0)
					kpif(0)=statkpfz("��������ǹ���������ɵļӹ�����", 0)
					kpif(1)=statkpfz("��������ӹ����ոĽ�������������", 0)
					kpif(2)=statkpfz("����ר�üо߳ɹ����ڼƻ���������", 0)
					kpif(3)=statkpfz("�½�������Ա���ܿ��˲��ϸ�", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(4)<-5 Then kpf(4)=-5
					ics(0)=statkpcs("�ϰ����빤���޹�", "", 0)
					ics(1)=statkpcs("ֵ�����,����", "", 0)
					ics(2)=statkpcs("���治��,�°����δ�ء���δ��", "", 0)
					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("ֵ�����,����", 0)
					kpif(2)=statkpfz("���治��,�°����δ�ء���δ��", 0)
					kpf(5)=kpif(0) + kpif(1) + kpif(2)
					If kpf(5)<-5 Then kpf(5)=-5
					ics(0)=statkpcs("�����ӷ���", "", 0)
					ics(1)=statkpcs("��������", "", 0)
					ics(2)=statkpcs("�������ⲻ��ʱ��������������", "", 0)
					ics(3)=statkpcs("�����е��������񲢻������", "", 0)
					kpif(0)=statkpfz("�����ӷ���", 0)
					kpif(1)=statkpfz("��������", 0)
					kpif(2)=statkpfz("�������ⲻ��ʱ��������������", 0)
					kpif(3)=statkpfz("�����е��������񲢻������", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(6)<-5 Then kpf(6)=-5
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>400 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(zrwfz/400 * 50),1)
				End If
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=FormatNumber(zrwxs+zzlxs,2)
			'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 9		'���Ա
					ics(0)=statkpcs("�ӳ�", "", 0)
					ics(1)=statkpcs("��ǰ", "", 0)
					kpif(0)=statkpfz("�ӳ�", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("��ǰ", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("����������ַ���", "", 0)
					ics(1)=statkpcs("��ͬ�����ظ�����2������", "", 0)
					ics(2)=statkpcs("����������ֱ���", "", 0)
					ics(3)=statkpcs("�ӹ����Լ췢�ֳ������", "", 0)
					ics(4)=statkpcs("��̷���ͼֽ���������", "", 0)
					kpif(0)=statkpfz("����������ַ���", 0)
					kpif(1)=statkpfz("��ͬ�����ظ�����2������", 0)
					kpif(2)=statkpfz("����������ֱ���", 0)
					kpif(3)=statkpfz("�ӹ����Լ췢�ֳ������", 0)
					kpif(4)=statkpfz("��̷���ͼֽ���������", 0)
					kpf(2)=kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)
					If kpf(2)<-20 Then kpf(2)=-20
					ics(0)=statkpcs("����ǳ��������豸�ӹ�����", "", 0)
					ics(1)=statkpcs("������ϸĽ���������", "", 0)
					ics(2)=statkpcs("�������ϴ���", "", 0)
					ics(3)=statkpcs("�½�������Ա���ܿ��˲��ϸ�", "", 0)
					kpif(0)=statkpfz("����ǳ��������豸�ӹ�����", 0)
					kpif(1)=statkpfz("������ϸĽ���������", 0)
					kpif(2)=statkpfz("�������ϴ���", 0)
					kpif(3)=statkpfz("�½�������Ա���ܿ��˲��ϸ�", 0)
					kpf(3)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(3)<-10 Then kpf(3)=-10
					If kpf(3)>10 Then kpf(3)=10
					ics(0)=statkpcs("�ϰ����빤���޹�", "", 0)
					ics(1)=statkpcs("�������������,����", "", 0)
					ics(2)=statkpcs("���治��,�°����δ�ء���δ��", "", 0)
					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("�������������,����", 0)
					kpif(2)=statkpfz("���治��,�°����δ�ء���δ��", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-2 Then kpf(4)=-2
					ics(0)=statkpcs("�����ӷ���", "", 0)
					ics(1)=statkpcs("�������ⲻ��ʱ", "", 0)
					ics(2)=statkpcs("��������", "", 0)
					kpif(0)=statkpfz("�����ӷ���", 0)
					kpif(1)=statkpfz("�������ⲻ��ʱ", 0)
					kpif(2)=statkpfz("��������", 0)
					kpf(5)=kpif(0) + kpif(1) + kpif(2)
					If kpf(5)<-5 Then kpf(5)=-5
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				If zrwfz>400 Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(zrwfz/400 * 50),1)
				End If
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=FormatNumber(zrwxs+zzlxs,2)
			'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 10		'������
			%>
<%Call diskpItem("��ƽ����ļ�δ����",5.0,2.5, -1)%>
<%Call diskpItem("���Խ����ļ�δ����",5.0,2.5, -1)%>
<%Call diskpItem("�����ļ�ǩ�𡢸��Ĳ�����",2.0,1.0, -1)%>
<%Call diskpItem("�ļ�����δ�浵",1.0,1.0, -1)%>
<%Call diskpItem("��Ŀ����������",3.0,1.0, -1)%>
<%Call diskpItem("��Ʒ�����¡���δ��ʱ�ϱ�",2.0,1.0, -1)%>
<%Call diskpItem("��Ʒ����ʵʩ�ƻ�ά�������",2.0,1.0, -1)%>
<%Call diskpItem("��ƹ淶ά��",6.0,2.0, 0)%>
<%Call diskpItem("��ͬ���������Ͳ����������ظ�",6.0,2.0, -1)%>
<%Call diskpItem("���ڵ�����Ϣ����",6.0,2.0, -1)%>
<%Call diskpItem("ģ�ߵ���δ�ϸ���",6.0,2.0, 0)%>
<%Call diskpItem("�ͻ���������Ͷ���뱧Թ",3.0,1.0, 0)%>
<%Call diskpItemM("���ԭ���������",8.0,0.4, "","���")%>
<%Call diskpItemM("���ԭ���������",6.0,0.4, "","���")%>
<%Call diskpItem("��׼�ṹ����δ�������",4.0,2.0, -1)%>
<%Call diskpItem("���ͷ���ǻ���ṹ���ٶȺ������鲻��",6.0,2.0, -1)%>
<%Call diskpItem("�ӿڼ����Ȱ彺��Ⱥ������鲻��",4.0,2.0, -1)%>
<%Call diskpItem("ͬ���β�Ʒ��ͬ��λģ����Ʋ�һ��",6.0,2.0, -1)%>
<%Call diskpItem("���׼ʱ�����",12.0,4.0, -1)%>
<%Call diskpItem("�������ⲻ��ʱ",1.0,1.0, -1)%>
<%Call diskpItem("����Э��",1.0,1.0, 0)%>
<%Call diskpItem("�����ձ��ܱ��±�",1.0,1.0, 0)%>
<%Call diskpItem("�·�����������Ŀ",1.0,1.0, 0)%>
<%Call diskpItem("����Ԥ����ʩ��ִ��",1.0,1.0, 0)%>
<%Call diskpItem("ʹ�÷���Ч�ļ�",1.0,1.0, 0)%>
<%Call diskpItem("�ֳ���������˾ͨ��",1.0,1.0, 0)%>
<%Call diskpItem("Ա����ѵ����������",1.0,1.0, 0)%>
<%	for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				zzlxs=FormatNumber(kpzf/100,2)
				zgkxs=FormatNumber(zrwxs+zzlxs,2)
			'	zyfgz=zjxgz*zgkxs*zbmxs+zjbgz

		Case 13		'�������Ա
				zzlxs=""
				zgkxs=""
				zyfgz=""

	End Select
	Erase kpf
End Function

Function ChkJs(str)
	'str ΪȨ��000001000000000
	ChkJs=0
	'If IsDebug Then ChkAble=True : Exit Function
	If Len(str)<15 Then Exit Function
	dim i
	For i=1 To Len(str)
		If Mid(str,i,1)=1 Then ChkJs=i : Exit For	'ֻȡÿ�˵���߽�ɫ,����ͬʱ���鳤����Ա,��ֻȡ�鳤
	Next
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

Function statkpjcs(kp_item,zrrjs,i)
	Dim tmpRs
	statkpjcs=0
	strSql="select * from [kp_jsb] where [kp_item] like '%"&kp_item&"%' and [kp_zrrjs] like '%"&zrrjs&"%' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	Call xjweb.Exec("",-1)
	Set tmpRs=Server.CreateObject("ADODB.RECORDSET")
	tmpRs.open strSql,Conn,1,3
		statkpjcs=tmpRs.RecordCount
	tmpRs.close
	set tmprs=nothing
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

Function statkpcs(kp_item, kp_zrrjs, i)
	Dim tmpRs
	statkpcs=0
	If kp_zrrjs="" Then
		Select Case i
			Case 0		'����Ա����ͳ��
				strSql=" [kp_jsb] where [kp_zrr]='"&zuser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'�����ν���ͳ��
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'���鳤����ͳ��
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 order by [kp_lsh]"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strSql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	else
				Select Case i
			Case 0		'����Ա����ͳ��
				strSql=" [kp_jsb] where [kp_zrr]='"&zuser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'�����ν���ͳ��
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'���鳤����ͳ��
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and [kp_zrrjs]='"&kp_zrrjs&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strSql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	End If
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
