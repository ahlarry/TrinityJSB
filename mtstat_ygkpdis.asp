<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'16:40 2017/2/22
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
	If InStr("1456",ChkJs(tmpAble))>0 Then		'�ж��ǲ�����Ա
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
		Case 1	'����
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd >������</td>
  <td class=ctd>�ճ�ά�������ϴ���<input name="basicwg" size="5" style="BACKGROUND-COLOR:transparent;BORDER-BOTTOM:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;BORDER-RIGHT:#ffffff 1px solid;BORDER-TOP:#ffffff 1px solid;COLOR:#00659c;HEIGHT:18px;border-color:#ffffff #ffffff #ffffff #ffffff;text-align:center;font-size:9pt" value=<%=TargetFZ%> >
  			<%if chkable(3) then%><input name="Chg_wg" type="submit" value="����"><%End If%>
  </td>
  <td class=ctd>50.0</td>
  <td class=ctd colspan=3 >����ϵͳ������</td>
  <%
'  	kpf(0)=round((iTotalFz/TargetFZ * 50),1)
  	kpf(0)=50
	%>
  <td class=ctd alt="<%="����:" & iTotalFz & "��"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=3>׼ȷ��</td>
  <td class=ctd>��������ƹ�Ӧ�ò���ʱ</td>
  <td class=ctd rowspan=3>40.0</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <%
		ics(0)=statkpcs("��������ƹ�Ӧ�ò���ʱ", "", 0)
		ics(1)=statkpcs("�������ϱ��ݲ���ʱ", "", 0)
		ics(2)=statkpcs("����Ȩ���趨����ȫ", "", 0)

		kpif(0)=statkpfz("��������ƹ�Ӧ�ò���ʱ", 0)
		kpif(1)=statkpfz("�������ϱ��ݲ���ʱ", 0)
		kpif(2)=statkpfz("����Ȩ���趨����ȫ", 0)

	kpf(1)=40+kpif(0)+kpif(1)+kpif(2)
	if kpf(1)<0 Then kpf(1)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=3><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>�������ϱ��ݲ���ʱ</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>����Ȩ���趨����ȫ</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>������</td>
  <td class=ctd>����̬�ȡ��Ͷ����ɿ۷�</td>
  <td class=ctd rowspan=2>10.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("����̬�ȡ��Ͷ����ɿ۷�", "", 0)
					ics(1)=statkpcs("����������ɲ���ʱ", "", 0)

					kpif(0)=statkpfz("����̬�ȡ��Ͷ����ɿ۷�", 0)
					kpif(1)=statkpfz("����������ɲ���ʱ", 0)

					kpf(2)=10 + kpif(0) + kpif(1)
					If kpf(2)<0 Then kpf(2)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>����������ɲ���ʱ</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				kpzf=kpf(0)+kpf(1)+kpf(2)
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
		Case 6	'����Ա
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd >������</td>
  <td class=ctd>���ԡ�����ͼ���֧�֣�<input name="basicwg" size="5" style="BACKGROUND-COLOR:transparent;BORDER-BOTTOM:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;BORDER-RIGHT:#ffffff 1px solid;BORDER-TOP:#ffffff 1px solid;COLOR:#00659c;HEIGHT:18px;border-color:#ffffff #ffffff #ffffff #ffffff;text-align:center;font-size:9pt" value=<%=TargetFZ%> >
  			<%if chkable(3) then%><input name="Chg_wg" type="submit" value="����"><%End If%>
  </td>
  <td class=ctd>50.0</td>
  <td class=ctd colspan=3 >����ϵͳ������</td>
  <%
'  	kpf(0)=round((iTotalFz/TargetFZ * 50),1)
  	kpf(0)=50
	%>
  <td class=ctd alt="<%="����:" & iTotalFz & "��"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>׼ʱ��</td>
  <td class=ctd>���������⴦����ʱ</td>
  <td class=ctd rowspan=2>10.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("�������⴦����ʱ", "", 0)
					ics(1)=statkpcs("���ڵ���δ׼ʱ���", "", 0)

					kpif(0)=statkpfz("�������⴦����ʱ", 0)
					kpif(1)=statkpfz("���ڵ���δ׼ʱ���", 0)

					kpf(1)=10+kpif(0) + kpif(1)
					if kpf(1)<0 Then kpf(1)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ڵ���δ׼ʱ���</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>׼ȷ��</td>
  <td class=ctd>���ԭ���������</td>
  <td class=ctd rowspan=4>30.0</td>
  <td class=ctd>4.0</td>
  <td class=ctd>��/��</td>
  <%
		ics(0)=statkpcs("������ԭ���������", "", 0)
		ics(1)=statkpcs("������ԭ���������", "", 0)
		ics(2)=statkpcs("���ԭ����ʧ��ǧԪ", "", 0)
		ics(3)=statkpcs("���ԭ���ⲿͶ��", "", 0)

		kpif(0)=statkpfz("������ԭ���������", 0)
		kpif(1)=statkpfz("������ԭ���������", 0)
		kpif(2)=statkpfz("���ԭ����ʧ��ǧԪ", 0)
		kpif(3)=statkpfz("���ԭ���ⲿͶ��", 0)

	kpf(2)=30+kpif(0)+kpif(1)+kpif(2)+kpif(3)
	if kpf(2)<0 Then kpf(2)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=4><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ���������(��)</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ����ʧ��ǧԪ</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/ǧԪ</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ���ⲿͶ��</td>
  <td class=ctd>3.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>������</td>
  <td class=ctd>����̬�ȡ��Ͷ����ɿ۷�</td>
  <td class=ctd rowspan=2>10.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("����̬�ȡ��Ͷ����ɿ۷�", "", 0)
					ics(1)=statkpcs("����������ɲ���ʱ", "", 0)

					kpif(0)=statkpfz("����̬�ȡ��Ͷ����ɿ۷�", 0)
					kpif(1)=statkpfz("����������ɲ���ʱ", 0)

					kpf(3)=10 + kpif(0) + kpif(1)
					If kpf(3)<0 Then kpf(3)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>����������ɲ���ʱ</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				kpzf=kpf(0)+kpf(1)+kpf(2)+kpf(3)
				%>
  <td class=ctd><%=kpzf%></td>
</tr>
<%
		Case else	'5����
			%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd >������</td>
  <td class=ctd>�������<input name="basicwg" size="5" style="BACKGROUND-COLOR:transparent;BORDER-BOTTOM:#ffffff 1px solid;BORDER-LEFT:#ffffff 1px solid;BORDER-RIGHT:#ffffff 1px solid;BORDER-TOP:#ffffff 1px solid;COLOR:#00659c;HEIGHT:18px;border-color:#ffffff #ffffff #ffffff #ffffff;text-align:center;font-size:9pt" value=<%=TargetFZ%> >
  			<%if chkable(3) then%><input name="Chg_wg" type="submit" value="����"><%End If%>
  </td>
  <td class=ctd>50.0</td>
  <td class=ctd colspan=3 >����ϵͳ������</td>
  <%
'  	kpf(0)=round((iTotalFz/TargetFZ * 50),1)
  	kpf(0)=50
	%>
  <td class=ctd alt="<%="����:" & iTotalFz & "��"%>"><%=kpf(0)%></td>
  <td class=ctd><%=kpf(0)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd >׼ʱ��</td>
  <td class=ctd>����ӳ�</td>
  <td class=ctd >10.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("����ӳ�", "", 0)

					kpif(0)=statkpfz("����ӳ�", 0)

					kpf(1)=10+kpif(0)
					if kpf(1)<0 Then kpf(1)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd><%=kpf(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=4>׼ȷ��</td>
  <td class=ctd>���ԭ���������</td>
  <td class=ctd rowspan=4>30.0</td>
  <td class=ctd>4.0</td>
  <td class=ctd>��/��</td>
  <%
  	if ChkJs(tmpAble)=4 Then
		ics(0)=statkpcs("���ԭ���������", "", tmpGroup)
		ics(1)=statkpcs("���ԭ���������", "", tmpGroup)
		kpif(0)=statkpfz("���ԭ���������", tmpGroup)
		kpif(1)=statkpfz("���ԭ���������", tmpGroup)
	Else
		ics(0)=statkpcs("���ԭ���������", "", 0)
		ics(1)=statkpcs("���ԭ���������", "", 0)
		kpif(0)=statkpfz("���ԭ���������", 0)
		kpif(1)=statkpfz("���ԭ���������", 0)
	End If
	ics(2)=statkpcs("���ԭ����ʧ��ǧԪ", "", 0)
	kpif(2)=statkpfz("���ԭ����ʧ��ǧԪ", 0)
	ics(3)=statkpcs("���ԭ���ⲿͶ��", "", 0)
	kpif(3)=statkpfz("���ԭ���ⲿͶ��", 0)

	kpf(2)=30+kpif(0)+kpif(1)+kpif(2)+kpif(3)
	if kpf(2)<0 Then kpf(2)=0
%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=4><%=kpf(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ���������(��)</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ����ʧ��ǧԪ</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/ǧԪ</td>
  <td class=ctd><%=ics(2)%></td>
  <td class=ctd><%=kpif(2)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>���ԭ���ⲿͶ��</td>
  <td class=ctd>3.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(3)%></td>
  <td class=ctd><%=kpif(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd rowspan=2>������</td>
  <td class=ctd>����̬�ȡ��Ͷ����ɿ۷�</td>
  <td class=ctd rowspan=2>10.0</td>
  <td class=ctd>2.0</td>
  <td class=ctd>��/��</td>
  <%
					ics(0)=statkpcs("����̬�ȡ��Ͷ����ɿ۷�", "", 0)
					ics(1)=statkpcs("����������ɲ���ʱ", "", 0)

					kpif(0)=statkpfz("����̬�ȡ��Ͷ����ɿ۷�", 0)
					kpif(1)=statkpfz("����������ɲ���ʱ", 0)

					kpf(3)=10 + kpif(0) + kpif(1)
					If kpf(3)<0 Then kpf(3)=0
				%>
  <td class=ctd><%=ics(0)%></td>
  <td class=ctd><%=kpif(0)%></td>
  <td class=ctd rowspan=2><%=kpf(3)%></td>
</tr>
<%icount=icount+1%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd>����������ɲ���ʱ</td>
  <td class=ctd>1.0</td>
  <td class=ctd>��/��</td>
  <td class=ctd><%=ics(1)%></td>
  <td class=ctd><%=kpif(1)%></td>
</tr>
<tr>
  <td class=rtd colspan=8>Total:</td>
  <%
				kpzf=kpf(0)+kpf(1)+kpf(2)+kpf(3)
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
	For i=1 To Len(str)
		If Mid(str,i,1)=1 Then ChkJs=i : Exit For	'ֻȡÿ�˵���߽�ɫ,����ͬʱ���鳤����Ա,��ֻȡ�鳤
	Next
End Function

Function statkpcs(kp_item, kp_zrrjs, i)
	Dim TmpRs
	statkpcs=0
	If kp_zrrjs="" Then
		Select Case i
			Case 0		'����Ա����ͳ��
				strSql=" [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
				statkpcs=xjweb.rscount(strSql)
			Case Else	'���鳤����ͳ��
				strSql="select [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 group by [kp_lsh]"
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
			Case Else	'���鳤����ͳ��
				strSql="select [kp_lsh] from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and [kp_zrrjs]='"&kp_zrrjs&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 group by [kp_lsh]"
				Set TmpRs=Server.CreateObject("adodb.recordset")
				TmpRs.open strsql,conn,1,3
				statkpcs=TmpRs.recordcount
				TmpRs.close
		End Select
	End If
End Function

Function statkpfz(kp_item, i)
	Dim ZzSql, ZzRs
	statkpfz=0 : ZzSql="" : ZzRs=""
	Dim tmpRs
	Select Case i
		Case 0		'����Ա����ͳ��
			strSql="select ([kp_uprice]*[kp_mul]) as kp_f from [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'���鳤����ͳ��
			strSql="select [kp_lsh],max([kp_uprice]*[kp_mul]*0.5) as kp_f from [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0 group by [kp_lsh]"
			ZzSql="select ([kp_uprice]*[kp_mul]*0.5) as kp_f from [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End Select

	Set tmpRs=xjweb.Exec(strSql, 1)
	do while not tmpRs.eof
		statkpfz=statkpfz + tmpRs("kp_f")
		tmpRs.movenext
	loop
	tmpRs.close
	set tmprs=nothing

	If i>0 Then
		Set ZzRs=xjweb.Exec(ZzSql, 1)
		do While not ZzRs.eof
			statkpfz=statkpfz + ZzRs("kp_f")
			ZzRs.movenext
		loop
		ZzRs.close
		set ZzRs=nothing
	End If
	statkpfz=round(statkpfz,2)
End Function
%>