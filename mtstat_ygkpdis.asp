<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'11:32 2007-4-10-���ڶ�
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="��ֵͳ�� �� �鿴������ֵͳ��"					'ҳ�������λ��( ��������� �� ���������)
strPage="mtstat"
xjweb.header()
Call TopTable()
'���������������ֵ			
Dim iyear, imonth, dtstart, dtend, struser, irwzf, iaddfz, ilxrwzf, icount
iyear = request("searchy")
imonth = request("searchm")
struser = request("searchuser")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)

dtend=cdate(iyear&"��"&imonth&"��1��")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
dtstart=cdate(iyear&"��"&imonth&"��1��")

'ͳ����
If struser = "" and chkable(5) Then struser = session("userName")
irwzf=0			'�ܷ�
ilxrwzf=0
iaddfz=0		'���ͷ�ֵ
icount=1		'������Ŀ��
'���¼���Ա����������
dim strChg_wg, strbasicwg
strChg_wg=request("Chg_wg")
strbasicwg=request("basicwg")
If strChg_wg="����" and chkable(3) then
	strSql="update [ims_user] set user_basicwage='"&strbasicwg&"' where user_name='"&struser&"'"
	call xjweb.Exec(strSql, 0)
End If
'���忼���õı���
	Dim kpf(30), kpif(10), ics(10), kpzf
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
  <Tr>
    <Td class=ctd height=300><%Call ygkpstatDisplay()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function SearchMantime()
%>
<table cellpadding=2 cellspacing=0>
  <form action=<%=request.servervariables("script_name")%> method=get>
    <tr>
      <td> ��ѡ��:
        <select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <%for i = year(now) - 3 to year(now)%>
          <option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        ��
        <select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <%for i = 1 to 12%>
          <option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        ��&nbsp;&nbsp;
        <select name="searchuser" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <option value=""></option>
          <%If chkable("1,2,3,4") Then%>
          <%for i = 0 to ubound(c_allstat)%>
          <option value="<%=c_allstat(i)%>" <%If struser = c_allstat(i) Then%>selected<%end If%>><%=c_allstat(i)%></option>
          <%next%>
          <%Else%>
          <option value="<%=session("userName")%>"><%=session("userName")%></option>
          <%end If%>
        </select>
        &nbsp;
        <input type="submit" value=" ѡ �� "></td>
    </tr>
  </form>
</table>
<%
End Function

Function ygkpstatDisplay()
	If struser="" Then Call TbTopic("��ѡ�������ѯ����Ա!") : Exit Function
	strSql="Select * from [ims_user] where [user_name]='"&struser&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then TbTopic("������ѡ���ѯ��Ա!") : Rs.Close : Exit Function
	Dim tmpGroup, tmpAble,TargetFZ
	tmpGroup=Rs("user_Group")
	tmpAble=Rs("user_Able")
	TargetFZ=Rs("user_basicwage")
	Rs.Close

	Dim iTotalFz, tmpCount			'�����ֵܷı���
	iTotalFz=0 : tmpCount=1
	If InStr("4568",ChkJs(tmpAble))>0 Then		'�ж��ǲ�����Ա�����Ա
		'1--�����ֵ
		strSql="select * from [mantime] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			irwzf=irwzf+Round(Rs("fz"),1)
			Rs.movenext
		Loop
		Rs.close
		'2---���������ֵ
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0"
		Set Rs=xjweb.Exec(strSql, 1)
		Do While Not Rs.eof
			ilxrwzf=ilxrwzf+Rs("zf")
			Rs.movenext
		Loop
		Rs.close
		'3---ͳ���ܷ�
		If Fix(ilxrwzf + irwzf)<(ilxrwzf + irwzf) Then
			iTotalFz=Fix(ilxrwzf + irwzf) + 1
		Else
			iTotalFz=Fix(ilxrwzf + irwzf)
		End If
	End If

	icount=1
	Call TbTopic(struser & " " & formatdatetime(dtstart,1) & " �� " & formatdatetime(dtend,1) & " ����ͳ��")
		%>
<table width="90%" cellpadding=2 cellspacing=0 class="xtable"  align="center" >
<form action=<%=request.servervariables("script_name")%> method="get">
<tr>
  <th class=th>id<input name="searchuser" type="hidden" value=<%=struser%>></th>
  <th class=th>������Ŀ</th>
  <th class=th>����ָ��</th>
  <th class=th>��Ԫ��(��)</th>
  <th class=th>�η�(��)</th>
  <th class=th>��λ</th>
  <th class=th>�ܴ���</th>
  <th class=th>����Ӧ�÷�</th>
  <th class=th>����ʵ�ʷ�</th>
</tr>
<%
	Select Case ChkJs(tmpAble)
		Case 6	'����Ա
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>������</td>
  <td class=ctd>ÿ�·�ֵ(����<%=TargetFZ%>��)</td>
  <td class=ctd>50.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <%
					If iTotalFz<TargetFZ Then
						kpf(0)=round((iTotalFz/TargetFZ * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-TargetFZ)/TargetFZ*50*1.25)),1)
					End If
				%>
  <td class=ctd alt="<%="����:" & iTotalFz & "��"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>�������</td>
  <td class=ctd>�ӳ�</td>
  <td class=ctd rowspan=2>10.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>��/��(����ƽ��)</td>
  <%
					ics(0)=statkpcs("�ӳ�", "", 0)
					ics(1)=statkpcs("��ǰ", "", 0)

					kpif(0)=statkpfz("�ӳ�", 0)
					If kpif(0)<-10 Then kpif(0)=-10
					kpif(1)=statkpfz("��ǰ", 0)
					If kpif(1)>10 Then kpif(1)=10
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>��ǰ</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>��/��(����ƽ��)</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>�����ƶ�</td>
  <td class=ctd>5����������δ����</td>
  <td class=ctd rowspan=4>8.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
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
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=4><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�������·�����ʱ</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�����鱨¼�벻��ʱ</td>
  <td class=ctd>1.5</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>����ͼֽǩ�𡢸��Ĳ�����</td>
  <td class=ctd>1.5</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>���Լ���</td>
  <td class=ctd>����ͼδ��ʱ�浵</td>
  <td class=ctd rowspan=3>20.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
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
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>������ԭ���������</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>������ԭ���������</td>
  <td class=ctd>4.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>�����֤</td>
  <td class=ctd>���Գ����������</td>
  <td class=ctd rowspan=2>&nbsp;</td>
  <td class=ctd>0.15</td>
  <td class=ctd>��ģ�߷�ֵ����������</td>
  <%
					ics(0)=statkpcs("���Գ����������", "", 0)
					ics(1)=statkpcs("�������ڶ��С����", "", 0)
					ics(2)=statkpcs("ģ�ߵ���δ�ϸ���", "", 0)

					kpif(0)=statkpfz("���Գ����������", 0)
					kpif(1)=statkpfz("�������ڶ��С����", 0)
					kpif(2)=statkpfz("ģ�ߵ���δ�ϸ���", 0)

					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-6 Then kpf(4)=-6
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�������ڶ��С����</td>
  <td class=ctd>0.15</td>
  <td class=ctd>��ģ�߷�ֵ�����ڴ���</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>ģ�ߵ���δ�ϸ���</td>
  <td class=ctd>6.0</td>
  <td class=ctd>3.0</td>
  <td class=ctd>��/��(ƽ��)</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>��������</td>
  <td class=ctd>�ϰ����빤���޹ص���</td>
  <td class=ctd rowspan=3>2.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("�ϰ����빤���޹�", "", 0)
					ics(1)=statkpcs("ֵ�����,����", "", 0)
					ics(2)=statkpcs("���治��,�°����δ�ء���δ��", "", 0)

					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("ֵ�����,����", 0)
					kpif(2)=statkpfz("���治��,�°����δ�ء���δ��", 0)

					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(6)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>ֵ�����,����</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���治��,�°����δ�ء���δ��</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>����̬��</td>
  <td class=ctd>�����ӷ���</td>
  <td class=ctd rowspan=2>4.0</td>
  <td class=ctd>4.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("�����ӷ���", "", 0)
					ics(1)=statkpcs("��������", "", 0)

					kpif(0)=statkpfz("�����ӷ���", 0)
					kpif(1)=statkpfz("��������", 0)

					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(7)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>��������</td>
  <td class=ctd>4.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				'����������Ĭ�ϵ�50��Ϊ������ֱ仯���仯,��������=�����/TargetFZ*50,��������50
				If iTotalFz>TargetFZ Then
					kpzf=kpzf+50
				else
					kpzf=round(kpzf+(iTotalFz/TargetFZ * 50),1)
				End If
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
		Case 8		'������Ա
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>��������</td>
  <td class=ctd>��Ʊ�׼���淶ά������ʱ</td>
  <td class=ctd rowspan=4>50.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("��Ʊ�׼���淶ά������ʱ", "", 0)
					ics(1)=statkpcs("���Է������⴦����ʱ", "", 0)
					ics(2)=statkpcs("Ӫ������֧�ֲ���ʱ", "", 0)
					ics(3)=statkpcs("������������", "", 0)

					kpif(0)=statkpfz("��Ʊ�׼���淶ά������ʱ", 0)
					kpif(1)=statkpfz("���Է������⴦����ʱ", 0)
					kpif(2)=statkpfz("Ӫ������֧�ֲ���ʱ", 0)
					kpif(3)=statkpfz("������������", 0)

					kpf(0)=50 + kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(0)<0 Then kpf(0)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>				
  <td class=ctd  rowspan=4><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���Է������⴦����ʱ</td>
  <td class=ctd>3.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>Ӫ������֧�ֲ���ʱ</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>��������׼ʱ��</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>��������</td>
  <td class=ctd>����ԭ�����ⲿͶ��</td>
  <td class=ctd rowspan=3>30.0</td>
  <td class=ctd>3.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("����ԭ�����ⲿͶ��", "", 0)
					ics(1)=statkpcs("��������δ���涨ִ��", "", 0)
					ics(2)=statkpcs("���ԭ��������ʧ��ǧԪ", "", 0)

					kpif(0)=statkpfz("����ԭ�����ⲿͶ��", 0)
					kpif(1)=statkpfz("��������δ���涨ִ��", 0)
					kpif(2)=statkpfz("���ԭ��������ʧ��ǧԪ", 0)

					kpf(1)=30 + kpif(0)+kpif(1)+kpif(2)
					If kpf(1)<0 Then kpf(1)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>��������δ���涨ִ��</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ��������ʧ��ǧԪ</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/ǧԪ</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				for i=0 to 1
					kpzf=kpzf+kpf(i)
				next
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%icount=1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=5>����ָ��</td>
  <td class=ctd>����Ľ�����ȡ�ó�Ч</td>
  <td class=ctd rowspan=5>20.0</td>
  <td class=ctd>1.0~5.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("����Ľ�����ȡ�ó�Ч", "", 0)
					ics(1)=statkpcs("�����е���������", "", 0)
					ics(2)=statkpcs("�ϰ����빤���޹�", "", 0)
					ics(3)=statkpcs("�����ӷ���", "", 0)
					ics(4)=statkpcs("5S�������", "", 0)

					kpif(0)=statkpfz("����Ľ�����ȡ�ó�Ч", 0)
					kpif(1)=statkpfz("�����е���������", 0)
					kpif(2)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(3)=statkpfz("�����ӷ���", 0)
					kpif(4)=statkpfz("5S�������", 0)

					kpf(2)=20 + kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4)
					If kpf(2)<0 Then kpf(2)=0					
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=5><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�����е���������</td>
  <td class=ctd>1.0~5.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�ϰ����빤���޹ص���</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�����ӷ���</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>5S�������</td>
  <td class=ctd>1.0~2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
    <%
				kpzf=kpf(2)
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
		Case Else	'��Ա
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>��������<br>(����50)</td>
  <td class=ctd>�������<input name="basicwg" size="5" style="BACKGROUND-COLOR:transparent;BORDER-BOTTOM:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;BORDER-RIGHT:#ffffff 1px solid;BORDER-TOP:#ffffff 1px solid;COLOR:#00659c;HEIGHT:18px;border-color:#ffffff #ffffff #ffffff #ffffff;text-align:center;font-size:9pt" value=<%=TargetFZ%> >
  			<%if chkable(3) then%><input name="Chg_wg" type="submit" value="����"><%End If%>
  </td>
  <td class=ctd>35.0</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
  <%
					If iTotalFz<TargetFZ Then
						kpf(0)=round((iTotalFz/TargetFZ * 35),1)
					Else
						kpf(0)=round((35+((iTotalFz-TargetFZ)/TargetFZ*35*1.25)),1)
					End If
				%>
  <td class=ctd alt="<%="����:" & iTotalFz & "��"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>��������׼ʱ��</td>
  <td class=ctd>5.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("������������", "", 0)
					kpif(0)=statkpfz("������������", 0)
					kpf(1)=kpif(0) + 5
					If kpf(1)<0 Then kpf(1)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>����׼ʱ��</td>
  <td class=ctd>10.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("�ӳ�", "", 0)
					kpif(0)=statkpfz("�ӳ�", 0)
					kpf(2)=kpif(0) + 10
					If kpf(2)<0 Then kpf(2)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=7>��������</td>
  <td class=ctd>���ԭ���������</td>
  <td class=ctd rowspan=7>30.0</td>
  <td class=ctd>4.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("���ԭ���������", "", 0)
					ics(1)=statkpcs("���ԭ���������", "", 0)
					ics(2)=statkpcs("���ԭ���������", "", 0)
					ics(3)=statkpcs("���ڵ������ڶ����", "", 0)
					ics(4)=statkpcs("���ڵ��Զ��ڶ����", "", 0)
					ics(5)=statkpcs("��������δ���涨ִ��", "", 0)
					ics(6)=statkpcs("���ԭ��������ʧ��ǧԪ", "", 0)

					kpif(0)=statkpfz("���ԭ���������", 0)
					kpif(1)=statkpfz("���ԭ���������", 0)
					kpif(2)=statkpfz("���ԭ���������", 0)
					kpif(3)=statkpfz("���ڵ������ڶ����", 0)
					kpif(4)=statkpfz("���ڵ��Զ��ڶ����", 0)
					kpif(5)=statkpfz("��������δ���涨ִ��", 0)
					kpif(6)=statkpfz("���ԭ��������ʧ��ǧԪ", 0)

					kpf(3)=30+kpif(0)+kpif(1)+kpif(2)+kpif(3)+kpif(4)+kpif(5)+kpif(6)
					if kpf(3)<0 Then kpf(3)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=7><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ���������</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ���������</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ڵ������ڶ����</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ڵ��Զ��ڶ����</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>��������δ���涨ִ��</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(5)%></td>
  <td class=ctd><%=kpif(5)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ��������ʧ��ǧԪ</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/ǧԪ</td>
  <td class=ctd><%=ics(6)%></td>
  <td class=ctd><%=kpif(6)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				for i=0 to 2
					kpzf=kpzf+kpf(i)
				next
				if kpzf>50 Then kpzf=50
				kpzf=kpzf+kpf(3)
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%icount=1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=5>����ָ��</td>
  <td class=ctd>����Ľ�����ȡ�ó�Ч</td>
  <td class=ctd rowspan=5>20.0</td>
  <td class=ctd>1.0~5.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("����Ľ�����ȡ�ó�Ч", "", 0)
					ics(1)=statkpcs("�����е���������", "", 0)
					ics(2)=statkpcs("�ϰ����빤���޹�", "", 0)
					ics(3)=statkpcs("�����ӷ���", "", 0)
					ics(4)=statkpcs("5S�������", "", 0)

					kpif(0)=statkpfz("����Ľ�����ȡ�ó�Ч", 0)
					kpif(1)=statkpfz("�����е���������", 0)
					kpif(2)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(3)=statkpfz("�����ӷ���", 0)
					kpif(4)=statkpfz("5S�������", 0)

					kpf(4)=20 + kpif(0) + kpif(1) + kpif(2) + kpif(3) + kpif(4)
					If kpf(4)<0 Then kpf(4)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=5><%=kpf(4)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�����е���������</td>
  <td class=ctd>1.0~5.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�ϰ����빤���޹ص���</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�����ӷ���</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>5S�������</td>
  <td class=ctd>1.0~2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(4)%></td>
  <td class=ctd><%=kpif(4)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				kpzf=kpf(4)
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
	End Select
	Erase kpf
End Function

Function ChkJs(str)
	'str ΪȨ��000001000000000
	ChkJs=0
	'If IsDebug Then ChkAble=True : Exit Function
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
	strSql="Select * from [ims_user] where mid(user_able,4,1)>0 and user_Group>0 and user_Group<4"
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
			strSql="select * from [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case -1		'�����ν���ͳ��
			strSql="select * from [kp_jsb] where [kp_item]='"&kp_item&"'  and [kp_kpr]<>" & struser & " and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'���鳤����ͳ��
			strSql="select * from [kp_jsb] where [kp_group]="&i&"  and [kp_kpr]<>" & struser & " and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
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
	Dim TmpRs
	statkpcs=0
	If kp_zrrjs="" Then
		Select Case i
			Case 0		'����Ա����ͳ��
				strSql=" [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'�����ν���ͳ��
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'���鳤����ͳ��
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 order by [kp_lsh]"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strsql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	else
		Select Case i
			Case 0		'����Ա����ͳ��
				strSql=" [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case -1		'�����ν���ͳ��
				strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'���鳤����ͳ��
				strSql="select distinct [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and [kp_zrrjs]='"&kp_zrrjs&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strsql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	End If
End Function

Function diskpItem(arg1,arg2,arg3,arg4)
	icount=icount+1
	dim tmpcs, tmpkpf
	%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>&nbsp;</td>
  <td class=ltd><%=arg1%></td>
  <td class=ctd><%=arg2%></td>
  <td class=ctd><%=arg3%></td>
  <td class=ctd>��/��Ŀ(����)</td>
  <%
					tmpcs=statkpcs(arg1, "", arg4)
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
				%>
  <td class=ctd><%=tmpcs%></td>
  <td class=ctd><%=tmpkpf%></td>
  <td class=ctd><%=kpf(icount-1)%></td>
</tr>
<%
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
	%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>&nbsp;</td>
  <td class=ltd><%=arg1%></td>
  <td class=ctd><%=arg2%>&nbsp;</td>
  <td class=ctd><%=arg3%></td>
  <td class=ctd>��/����</td>
  <%
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If arg2<>"" Then
						If kpf(icount-1)<arg2*-1 Then
							 kpf(icount-1)=arg2*-1
						End If
					End If
					'�鳤���ޡ�����������
				%>
  <td class=ctd><%=tmpcs%></td>
  <td class=ctd><%=tmpkpf%></td>
  <td class=ctd><%=kpf(icount-1)%></td>
</tr>
<%
	Rs.Close
End Function
%>
