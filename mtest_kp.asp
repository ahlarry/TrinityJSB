<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="ģ�ߵ��� �� ���Կ����б�"
strPage="mtest"
'Call FileInc(0, "js/ftask.js")
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

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchMantime()%>
		</Td></Tr>
		<Tr><Td class=ctd height=300>
			<%Call mtestList()%>
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
			<select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value);'>
				<%for i = year(now) - 3 to year(now)%>
					<option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>��
			<select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value);'>
				<%for i = 1 to 12%>
					<option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>��&nbsp;&nbsp;
			<input type="submit" value=" ѡ �� ">
			</td>
		</tr>
		</form>
	</table>
<%
End Function

Function mtestList()
	Call TbTopic("���������Ϣ")
	Dim Tmplsh,hgs,hgf,hgz,cts,ctf,ctz,jts,jtf,jtz,jys,jyf,jyz,lcs,lcf,lcz
	Tmplsh="" : hgs=0 : hgf=0 : hgz=0 : cts=0 : ctf=0 : ctz=0 : jts=0 : jtf=0 : jtz=0 : jys=0 : jyf=0 : jyz=0 : lcs=0 : lcf=0 : lcz=0
	strSql="select * from [mantime] where zrr='TT����Ա' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Do While Not Rs.eof
		If Tmplsh<>Rs("lsh") Then
			select case Mid(Rs("rwlr"),3)
				case "���ڳ���"
					cts=cts+1
					If Rs("jc")>1 Then ctz=ctz+1
					If Rs("jc")<1 Then ctf=ctf+1
				case "���⾫��"
					jts=jts+1
					If Rs("jc")>1 Then jtz=jtz+1
					If Rs("jc")<1 Then jtf=jtf+1
				case "Ԥ���ջ����"
					jys=jys+1
					If Rs("jc")>1 Then jyz=jyz+1
					If Rs("jc")<1 Then jyf=jyf+1
				case "��������"
					lcs=lcs+1
					If Rs("jc")>1 Then lcz=lcz+1
					If Rs("jc")<1 Then lcf=lcf+1
				case else
					hgs=hgs+1
					If Rs("jc")>1 Then hgz=hgz+1
					If Rs("jc")<1 Then hgf=hgf+1
			end select
			Tmplsh=Rs("lsh")
		End If
		Rs.movenext
	Loop
	Rs.Close
%>
<table width="80%" cellpadding=2 cellspacing=0 class="xtable" align="center">
  <tr>
    <th class=th>&nbsp;</th>
    <th class=th>���Ժϸ�</th>
    <th class=th>���ڳ���</th>
    <th class=th>���⾫��</th>
    <th class=th>���ռ���</th>
    <th class=th>��������</th>
    <th class=th>�ϼ���Ч����</th>
  </tr>
  <tr>
    <td class=ctd>�������˴���</th>
    <td class=ctd><%=hgf%></td>
    <td class=ctd><%=ctf%></td>
    <td class=ctd><%=jtf%></td>
    <td class=ctd><%=jyf%></td>
    <td class=ctd><%=lcf%></td>
    <td class=ctd><%=hgf+ctf+jtf+jyf+lcf%></td>
  </tr>
  <tr>
    <td class=ctd>���ڶ���˴���</th>
    <td class=ctd><%=hgz%></td>
    <td class=ctd><%=ctz%></td>
    <td class=ctd><%=jtz%></td>
    <td class=ctd><%=jyz%></td>
    <td class=ctd><%=lcz%></td>
    <td class=ctd><%=hgz+ctz+jtz+jyz+lcz%></td>
  </tr>
  <tr>
    <td class=ctd>�ϼ�(�������˷�Χ��ģ��)</td>
    <td class=ctd><%=hgs%></td>
    <td class=ctd><%=cts%></td>
    <td class=ctd><%=jts%></td>
    <td class=ctd><%=jys%></td>
    <td class=ctd><%=lcs%></td>
    <td class=ctd><%=hgs+cts+jts+jys+lcs%></td>
  </tr>
</table>
<%Call TbTopic("��ֵͳ��")
	Dim sjjl,sjcf
	sjjl=0 : sjcf=0
  	strSql="select * from [mantime] where rwlr like '%���Ժϸ�(%' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Do While Not Rs.eof
		If Rs("fz")>0 Then
			sjjl=sjjl+Rs("fz")
		else
			sjcf=sjcf+Rs("fz")
		End If
		Rs.movenext
	Loop
	Rs.Close

	Dim sjzzjl,sjzzcf,tszzjl,tszzcf,xxzzjl,xxzzcf,gfwhjl,gfwhcf,jsdbjl,jsdbcf
	sjzzjl=0 : sjzzcf=0 : tszzjl=0 : tszzcf=0 : xxzzjl=0
	xxzzcf=0 : gfwhjl=0 : gfwhcf=0 : jsdbjl=0 : jsdbcf=0
  	strSql="select * from [kp_jsb] where kp_item like '%�ڶ����%' and datediff('d',kp_time,'"&dtstart&"')<=0 and datediff('d',kp_time,'"&dtend&"')>=0"
	Set Rs=xjweb.Exec(strSql, 1)
	Do While Not Rs.eof
		select case Rs("kp_zrrjs")
			case "����鳤"
				If Rs("kp_item")="�������ڶ����" Then
					sjzzjl=sjzzjl+Rs("kp_uprice")
				else
					sjzzcf=sjzzcf-Rs("kp_uprice")
				End If
			case "�����鳤"
				If Rs("kp_item")="�������ڶ����" Then
					tszzjl=tszzjl+Rs("kp_uprice")
				else
					tszzcf=tszzcf-Rs("kp_uprice")
				End If
			case "��Ϣ�鳤"
				If Rs("kp_item")="�������ڶ����" Then
					xxzzjl=xxzzjl+Rs("kp_uprice")
				else
					xxzzcf=xxzzcf-Rs("kp_uprice")
				End If
			case "�淶ά��"
				If Rs("kp_item")="�������ڶ����" Then
					gfwhjl=gfwhjl+Rs("kp_uprice")
				else
					gfwhcf=gfwhcf-Rs("kp_uprice")
				End If
			case "��������"
				If Rs("kp_item")="�������ڶ����" Then
					jsdbjl=jsdbjl+Rs("kp_uprice")
				else
					jsdbcf=jsdbcf-Rs("kp_uprice")
				End If
		end select
		Rs.movenext
	Loop
	Rs.Close

	sjzzjl=Round(sjzzjl,1)
	sjzzcf=Round(sjzzcf,1)
	tszzjl=Round(tszzjl,1)
	tszzcf=Round(tszzcf,1)
	xxzzjl=Round(xxzzjl,1)
	xxzzcf=Round(xxzzcf,1)
	gfwhjl=Round(gfwhjl,1)
	gfwhcf=Round(gfwhcf,1)
	jsdbjl=Round(jsdbjl,1)
	jsdbcf=Round(jsdbcf,1)
%>
<table width="80%" cellpadding=2 cellspacing=0 class="xtable" align="center">
  <tr>
    <th class=th>&nbsp;</th>
    <th class=th>�����Ա</th>
    <th class=th>����鳤</th>
    <th class=th>�����鳤</th>
    <th class=th>��Ϣ�鳤</th>
    <th class=th>�淶ά��</th>
    <th class=th>��������</th>
  </tr>
  <tr>
    <td class=ctd>����</td>
    <td class=ctd><%=sjjl%></td>
    <td class=ctd><%=sjzzjl%></td>
    <td class=ctd><%=tszzjl%></td>
    <td class=ctd><%=xxzzjl%></td>
    <td class=ctd><%=gfwhjl%></td>
    <td class=ctd><%=jsdbjl%></td>
  </tr>
  <tr>
    <td class=ctd>����</th>
    <td class=ctd><%=sjcf%></td>
    <td class=ctd><%=sjzzcf%></td>
    <td class=ctd><%=tszzcf%></td>
    <td class=ctd><%=xxzzcf%></td>
    <td class=ctd><%=gfwhcf%></td>
    <td class=ctd><%=jsdbcf%></td>
  </tr>
  <tr>
    <td class=ctd>�ϼ�</th>
    <td class=ctd><%=sjjl+sjcf%></td>
    <td class=ctd><%=sjzzjl+sjzzcf%></td>
    <td class=ctd><%=tszzjl+tszzcf%></td>
    <td class=ctd><%=xxzzjl+xxzzcf%></td>
    <td class=ctd><%=gfwhjl+gfwhcf%></td>
    <td class=ctd><%=jsdbjl+jsdbcf%></td>
  </tr>
</table>
<%
end function
%>