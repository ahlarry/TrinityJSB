<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="��ֵͳ�� �� Test06"					'ҳ�������λ��( ��������� �� ���������)
strPage="mtstat"
'Call FileInc(0, "js/login.js")
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

'���忼���õı���
	Dim kpf(30), kpif(5), ics(5), kpzf
	kpzf=0
	for i=0 to 29
		kpf(i)=0
	next
	for i=0 to 4
		kpif(i)=0
	next
	for i=0 to 4
		ics(i)=0
	next

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchMantime()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call ygkpstatDisplay()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub
	
Function SearchMantime()
%>
	<table cellpadding=2 cellspacing=0>
		<form action=<%=request.servervariables("script_name")%> method=get>
		<tr>
			<td>
			��ѡ��:
			<select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
				<%for i = year(now) - 3 to year(now)%>
					<option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>��
			<select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
				<%for i = 1 to 12%>
					<option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>��&nbsp;&nbsp;

			<select name="searchuser" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
				<option value=""></option>
				<%If chkable("1,2,3,4") Then%>
					<%for i = 0 to ubound(c_allstat)%>
						<option value="<%=c_allstat(i)%>" <%If struser = c_allstat(i) Then%>selected<%end If%>><%=c_allstat(i)%></option>
					<%next%>
				<%Else%>
					<option value="<%=session("userName")%>"><%=session("userName")%></option>
				<%end If%>
			</select>&nbsp;
			<input type="submit" value=" ѡ �� ">
			</td>
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
	Dim tmpGroup, tmpAble
	tmpGroup=Rs("user_Group")
	tmpAble=Rs("user_Able")
	Rs.Close

	Dim iTotalFz			'�����ֵܷı���

	If ChkJs(tmpAble)=5 Or ChkJs(tmpAble)=6 Then		'�ж��ǲ�����Ա�����Ա
		'�������Ա�����Ա�Ļ�������ͳ�������ֵ
		'ͳ�Ʒ�ֵ
		'1--�����ֵ
		strSql="select a.*, b.* ,a.lsh as lsh, a.rwlr as rwlr from [mantime] a, [mtask] b where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 and a.lsh=b.lsh order by jssj desc, a.lsh desc"
		Set Rs=xjweb.Exec(strSql, 1)
		Dim itmpLsh, bJc		'������ʱ����
		itmpLsh="" : bJc=True
		If Not(Rs.eof or Rs.bof) Then 
			Do While Not Rs.eof
				If itmpLsh<>Rs("lsh") Then
					itmpLsh=Rs("lsh")
					bJc=True
				Else
					bJc=False
				End If
				irwzf=irwzf+Round(Rs("fz"),1)
				If bJc Then iaddfz=iaddfz+Rs("jc")
				Rs.movenext
			Loop
		End If
		Rs.close

		'2---���������ֵ
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by jssj desc"
		Set Rs=xjweb.Exec(strSql, 1)
		If Not(Rs.eof or Rs.bof) Then 
			Do While Not Rs.eof
				ilxrwzf=ilxrwzf+Rs("zf")
				Rs.movenext
			Loop
		End If
		Rs.close

		'3---ͳ���ܷ�
		
		If Fix(ilxrwzf + irwzf + iaddfz)<(ilxrwzf + irwzf + iaddfz) Then
			iTotalFz=Fix(ilxrwzf + irwzf + iaddfz) + 1
		Else
			iTotalFz=Fix(ilxrwzf + irwzf + iaddfz)
		End If

	End If		'�ж��ǲ�����Ա�͵���Ա����

	
	
	icount=1

	Call TbTopic(struser & " " & formatdatetime(dtstart,1) & " �� " & formatdatetime(dtend,1) & " ����ͳ��")
		%>
		<table width="90%" cellpadding=2 cellspacing=0 class="xtable">
	<%
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
			<%Call diskpItemM("���ԭ���������",8.0,0.4, -1, 2)%>
			<%Call diskpItemM("���ԭ���������",6.0,0.4, -1, 2)%>
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
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			</tr>
			<%
		Case 4	'�鳤
			%>
			<%Call diskpItem("�ļ�δ��ָ��·������",5.0,2.5, 0)%>
			<%Call diskpItem("�����ļ�ǩ�𡢸��Ĳ�����",2.0,1.0, tmpGroup)%>
			<%Call diskpItem("��Ʒ�����¼ƻ�����δ���",6.0,2.0, tmpGroup)%>
			<%Call diskpItem("ģ����������󲻳�ַ���",6.0,3.0, 0)%>
			<%Call diskpItem("��׼�ṹ����δ�������",4.0,2.0, tmpGroup)%>
			<%Call diskpItem("��ͬ���������Ͳ����������ظ�",6.0,2.0, 0)%>
			<%Call diskpItem("���ڵ�����Ϣ����",4.0,2.0, 0)%>
			<%Call diskpItemM("���ԭ���������",6.0,0.6, tmpGroup, 2)%>
			<%Call diskpItemM("���ԭ���������",8.0,0.6, tmpGroup, 2)%>
			<%Call diskpItemM("���и����1:1ͼ",4.0,0.6, tmpGroup, 2)%>
			<%Call diskpItem("�ṹ���������֯����ʱ",4.0,1.0, 0)%>
			<%Call diskpItem("���ͷ���ǻ���ṹ���ٶȺ������鲻��",8.0,2.0, 0)%>
			<%Call diskpItem("�ӿڼ����Ȱ彺��Ⱥ������鲻��",4.0,2.0, 0)%>
			<%Call diskpItem("ͬ���β�Ʒ��ͬ��λģ����Ʋ�һ��",6.0,3.0, 0)%>
			<%Call diskpItem("���׼ʱ�����",12.0,4.0, 0)%>
			<%Call diskpItem("�������ⲻ��ʱ",2.0,1.0, 0)%>
			<%Call diskpItem("�����ܱ�",2.0,1.0, 0)%>
			<%Call diskpItem("С������",1.0,1.0, 0)%>
			<%Call diskpItem("�������������ƻ�������",2.0,1.0, 0)%>
			<%Call diskpItem("�������Ԥ����ʩ��ִ��",2.0,1.0, 0)%>
			<%Call diskpItem("Ա����ѵ����������",1.0,1.0, 0)%>
			<%Call diskpItem("�ֳ���������",1.0,1.0, 0)%>
			<%Call diskpItem("��ԱͶ��",2.0,1.0, 0)%>
			<%Call diskpItem("������������������",2.0,1.0, 0)%>
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			</tr>
			<%
		Case 6	'����Ա
					If iTotalFz<300 Then
						kpf(0)=round((iTotalFz/300 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-300)*0.25)),1)
					End If
				%>
				<td class=ctd alt="<%="������:" & iTotalFz & "��"%>"><%=kpf(0)%></td>
				<%
					ics(0)=statkpcs("�ӳ�", 0)
					ics(1)=statkpcs("��ǰ", 0)					
					kpif(0)=statkpfz("�ӳ�", 0)
					kpif(1)=statkpfz("��ǰ", 0)
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("5����������δ����", 0)
					ics(1)=statkpcs("�������·�����ʱ", 0)
					ics(2)=statkpcs("�����鱨¼�벻��ʱ", 0)
					ics(3)=statkpcs("����ͼֽǩ�𡢸��Ĳ�����", 0)					
					kpif(0)=statkpfz("5����������δ����", 0)
					kpif(1)=statkpfz("�������·�����ʱ", 0)
					kpif(2)=statkpfz("�����鱨¼�벻��ʱ", 0)
					kpif(3)=statkpfz("����ͼֽǩ�𡢸��Ĳ�����", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2) + kpif(3)
					If kpf(2)<-8 Then kpf(2)=-8
					ics(0)=statkpcs("����ͼδ��ʱ�浵", 0)
					ics(1)=statkpcs("������ԭ���������", 0)
					ics(2)=statkpcs("������ԭ���������", 0)
					kpif(0)=statkpfz("����ͼδ��ʱ�浵", 0)
					kpif(1)=statkpfz("������ԭ���������", 0)
					kpif(2)=statkpfz("������ԭ���������", 0)
					kpf(3)=kpif(0)+kpif(1)+kpif(2)
					If kpf(3)<-20 Then kpf(3)=-20
					ics(0)=statkpcs("�ڶ�����Ʒ�ϸ�", 0)
					ics(1)=statkpcs("��������Ʒ�ϸ�", 0)
					ics(2)=statkpcs("ģ�ߵ���δ�ϸ���", 0)
					kpif(0)=statkpfz("�ڶ�����Ʒ�ϸ�", 0)
					kpif(1)=statkpfz("��������Ʒ�ϸ�", 0)
					kpif(2)=statkpfz("ģ�ߵ���δ�ϸ���", 0)
					kpf(4)=kpif(0) + kpif(1) + kpif(2)
					If kpf(4)<-6 Then kpf(4)=-6
					ics(0)=statkpcs("�ϰ����빤���޹�", 0)
					ics(1)=statkpcs("ֵ�����,����", 0)
					ics(2)=statkpcs("���治��,�°����δ�ء���δ��", 0)
					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("ֵ�����,����", 0)
					kpif(2)=statkpfz("���治��,�°����δ�ء���δ��", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
					ics(0)=statkpcs("�����ӷ���", 0)
					ics(1)=statkpcs("��������", 0)
					kpif(0)=statkpfz("�����ӷ���", 0)
					kpif(1)=statkpfz("��������", 0)
					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				%>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+50
				%>
				<td class=ctd><%=kpzf%></td>
			<%
		Case 8		'����ָ��������
			%>
			<%Call diskpItem("��ƹ淶����׼����ά��",15.0,5.0, 0)%>
			<%Call diskpItem("ģ����������󲻳�ַ���",9.0,3.0, -1)%>
			<%Call diskpItem("��׼�ṹ����δ�������",16.0,4.0, -1)%>
			<%Call diskpItem("ͬ���β�Ʒ��ͬ��λģ����Ʋ�һ��",12.0,4.0, -1)%>
			<%Call diskpItem("��ͬ���������Ͳ����������ظ�",16.0,4.0, -1)%>
			<%Call diskpItemM("���ԭ���������",8.0,2.0, -1, 2)%>
			<%Call diskpItemM("���ԭ���������",12.0,4.0, -1, 2)%>
			<%Call diskpItem("���ͷ���ǻ���ṹ���ٶȺ������鲻��",8.0,2.0, -1)%>
			<%Call diskpItem("�ӿڼ����Ȱ彺��Ⱥ������鲻��",4.0,2.0, -1)%>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			<%
		Case 5	'��Ա
					If iTotalFz<300 Then
						kpf(0)=round((iTotalFz/300 * 50),1)
					Else
						kpf(0)=round((50+((iTotalFz-300)*0.25)),1)
					End If
				%>
				<td class=ctd alt="<%="������:" & iTotalFz & "��"%>"><%=kpf(0)%></td>
				<%
					ics(0)=statkpcs("�ӳ�", 0)
					ics(1)=statkpcs("��ǰ", 0)
					kpif(0)=statkpfz("�ӳ�", 0)
					kpif(1)=statkpfz("��ǰ", 0)
					kpf(1)=kpif(0) + kpif(1)
					If kpf(1)>10 Then kpf(1)=10
					If kpf(1)<-10 Then kpf(1)=-10
					ics(0)=statkpcs("��Ʒ�����¼ƻ�����δ���", 0)
					ics(1)=statkpcs("��Ʒ�����¡���δ��ʱ�ϱ�", 0)
					ics(2)=statkpcs("��Ʒ����ʵʩ�ƻ�ά�������", 0)
					kpif(0)=statkpfz("��Ʒ�����¼ƻ�����δ���", 0)
					kpif(1)=statkpfz("��Ʒ�����¡���δ��ʱ�ϱ�", 0)
					kpif(2)=statkpfz("��Ʒ����ʵʩ�ƻ�ά�������", 0)
					kpf(2)=kpif(0) + kpif(1) + kpif(2)
					If kpf(2)<-8 Then kpf(2)=-8
					ics(0)=statkpcs("�����ļ�ǩ�𡢸��Ĳ�����", 0)
					kpif(0)=statkpfz("�����ļ�ǩ�𡢸��Ĳ�����", 0)
					kpf(3)=kpif(0)
					If kpf(3)<-4 Then kpf(3)=-4
					ics(0)=statkpcs("��׼�ṹ����δ�������", 0)
					ics(1)=statkpcs("���и����1:1ͼ", 0)
					ics(2)=statkpcs("���ԭ���������", 0)
					ics(3)=statkpcs("���ԭ���������", 0)
					kpif(0)=statkpfz("��׼�ṹ����δ�������", 0)
					kpif(1)=statkpfz("���и����1:1ͼ", 0)
					kpif(2)=statkpfz("���ԭ���������", 0)
					kpif(3)=statkpfz("���ԭ���������", 0)
					kpf(4)=kpif(0)+kpif(1)+kpif(2)+kpif(3)
					If kpf(4)<-22 Then kpf(4)=-22
					ics(0)=statkpcs("��һ����Ʒ�ϸ�", 0)
					ics(1)=statkpcs("�ڶ�����Ʒ�ϸ�", 0)
					kpif(0)=statkpfz("��һ����Ʒ�ϸ�", 0)
					kpif(1)=statkpfz("�ڶ�����Ʒ�ϸ�", 0)
					kpf(5)=kpif(0) + kpif(1)
					ics(0)=statkpcs("�ϰ����빤���޹�", 0)
					ics(1)=statkpcs("ֵ�����,����", 0)
					ics(2)=statkpcs("���治��,�°����δ�ء���δ��", 0)
					kpif(0)=statkpfz("�ϰ����빤���޹�", 0)
					kpif(1)=statkpfz("ֵ�����,����", 0)
					kpif(2)=statkpfz("���治��,�°����δ�ء���δ��", 0)
					kpf(6)=kpif(0) + kpif(1) + kpif(2)
					If kpf(6)<-2 Then kpf(6)=-2
					ics(0)=statkpcs("�����ӷ���", 0)
					ics(1)=statkpcs("��������", 0)
					kpif(0)=statkpfz("�����ӷ���", 0)
					kpif(1)=statkpfz("��������", 0)
					kpf(7)=kpif(0) + kpif(1)
					If kpf(7)<-4 Then kpf(7)=-4
				%>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+50
				%>
				<td class=ctd><%=kpzf%></td>			
			<%
		Case 9			'��Ŀ��������
			%>
			<%Call diskpItem("�ѽ�����Ŀδ����",10.0,2.5, 0)%>
			<%Call diskpItem("��Ŀ�ļ�ǩ�𲻱�׼",15.0,3.0, 0)%>
			<%Call diskpItem("��Ŀ�����±���ʱ����",15.0,3.0, 0)%>
			<%Call diskpItem("��Ŀ�±�����׼ʱ���",15.0,3.0, 0)%>
			<%Call diskpItem("��Ŀ�����ƻ��������",10.0,5.0, 0)%>
			<%Call diskpItem("��Ŀ��������������",20.0,4.0, 0)%>
			<%Call diskpItem("��Ŀ��������δ���",15.0,5.0, 0)%>
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
				%>
				<td class=ctd><%=kpzf%></td>
			</tr>
			<%
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
			<%Call diskpItemM("���ԭ���������",8.0,1.0, -1,2)%>
			<%Call diskpItemM("���ԭ���������",6.0,3.0, -1,2)%>
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
			<tr>
				<td class=rtd colspan=8>Total:</td>
				<%
				for i=0 to 29
					kpzf=kpzf+kpf(i)
				next
				kpzf=kpzf+100
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

Function statkpcs(kp_item, i)
	statkpcs=0
	Select Case i
		Case 0		'����Ա����ͳ��
			strSql=" [kp_jsb] where [kp_zrr]='"&struser&"' and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case -1		'�����ν���ͳ��
			strSql=" [kp_jsb] where [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
		Case Else	'���鳤����ͳ��
			strSql=" [kp_jsb] where [kp_group]="&i&" and [kp_item]='"&kp_item&"' and datediff('d',[kp_time],'"&dtstart&"')<=0 and datediff('d',[kp_time],'"&dtend&"')>=0"
	End Select
	statkpcs=xjweb.rscount(strSql)
End Function


Function diskpItem(arg1,arg2,arg3, arg4)
	dim tmpcs, tmpkpf
					tmpcs=statkpcs(arg1, arg4)
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
End Function

Function diskpItemM(arg1,arg2,arg3, arg4, arg5)
	icount=icount+1
	dim tmpcs, tmpkpf
					If Instr(arg1,"���ԭ���������")>0 or Instr(arg1,"���и����1:1ͼ") Then arg3=2*arg3
					If Instr(arg1,"���ԭ���������")>0  Then arg3=4*arg3
					tmpcs=Int(statkpcs(arg1, arg4)/arg5)
					tmpkpf=tmpcs*arg3*-1
					kpf(icount-1)=tmpkpf
					If kpf(icount-1)<arg2*-1 Then kpf(icount-1)=arg2*-1
End Function
%>