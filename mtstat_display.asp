<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="��ֵͳ�� �� �鿴�����ֵͳ��"					'ҳ�������λ��( ��������� �� ���������)
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

'��ʼ����
'����ʱ����ʱͣ��
'dtend=cdate(iyear&"��"&imonth&"��10��")
'dtstart=dateadd("m", -1, dtend)		'�ϸ���
'dtstart=dateadd("d", 1, dtstart)		'�ϸ���11��

dtend=cdate(iyear&"��"&imonth&"��1��")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
'dtstart=dateadd("m", -1, dtend)		'�ϸ���
dtstart=cdate(iyear&"��"&imonth&"��1��")


'dtstart=dateadd("d", 6, dtstart)		'�ϸ���11��
'ͳ����
If struser = "" and chkable(5) Then struser = session("userName")
irwzf=0			'�ܷ�
ilxrwzf=0
iaddfz=0		'���ͷ�ֵ
icount=1		'������Ŀ��

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>"  align="center">
  <Tr>
    <Td class=ctd><%Call SearchMantime()%></td>
  </tr>
  <Tr>
    <Td class=ctd height=300><%Call mtstatDisplay()%>
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

Function mtstatDisplay()
	If struser="" Then
		Call TbTopic("��ѡ�������ѯ����Ա!")
	Else
		Call TbTopic(struser & " " & formatdatetime(dtstart,1) & " �� " & formatdatetime(dtend,1) & " ��ֵͳ��")
		%>
<table width="80%" cellpadding=2 cellspacing=0 class="xtable"  align="center">
  <tr>
    <th class=th>id</th>
    <th class=th>��ˮ��</th>
    <th class=th>��������</th>
    <th class=th>������</th>
    <th class=th>�������</th>
    <th class=th>�����ֵ</th>
    <th class=th>ϵ��</th>
    <th class=th>����</th>
  </tr>
  <%
			call mtask_mt()
			call ftask_mt()
			call total_mt()
		%>
</table>
<%
	End If
End Function

function mtask_mt()		'��������ֵͳ��
	strSql="select a.*, b.* ,a.lsh as lsh, a.rwlr as rwlr from [mantime] a, [mtask] b where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 and a.lsh=b.lsh order by jssj desc, a.lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Dim itmpLsh, bJc		'������ʱ����
	itmpLsh="" : bJc=True
	Do While Not Rs.eof
	%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd title="<p>������: <%=Rs("ddh")%></p>��ˮ��: <%=Rs("lsh")%><br>�ͻ�����: <%=Rs("dwmc")%><br>��������: <%=Rs("dmmc")%><br>"><a href="mtask_display.asp?s_lsh=<%=Rs("lsh")%>"><%=Rs("lsh")%></a></td>
  <td class=ctd><%=Rs("rwlr")%></td>
  <td class=ctd><%=Rs("zrr")%></td>
  <td class=ctd alt="�ƻ�����ʱ��:<%=xjDate(Rs("jhjssj"),1)%>"><%=xjDate(Rs("jssj"),1)%></td>
  <td class=ctd><%=Round(Rs("fz"),1)%></td>
  <td class=ctd><%If Rs("jc")>0 Then Response.Write(Rs("jc")) end If%>
    &nbsp; </td>
  <td class=ctd><%if (InStr(Rs("rwlr"),"���Ժϸ�(")>0 and chkable(3)) Then%>
    <input type=button id=<%=Rs("a.id")%> value="�޸�" onClick="changesjf(this.id)">
    <%End If%>
    &nbsp;</td>
</tr>
<%
		icount = icount + 1
		irwzf=irwzf+Round(Rs("fz"),1)
'		If bJc Then iaddfz=iaddfz+Rs("jc")
		Rs.movenext
	loop
	%>
<tr>
  <td class=rtd colspan=5>�����ܷ�:</td>
  <td class=ctd colspan=3><b><%=irwzf%></b></td>
</tr>
<%
	Rs.close
end function

function ftask_mt()		'���������ֵͳ��
		strSql="select * from [ftask] where zrr='"&struser&"' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by jssj desc"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.eof or Rs.bof Then
		'response.write("<tr><td class="ctd" colspan=7>û���κ���������</td></tr>")
	Else
		Do While Not Rs.eof
		%>
<tr>
  <td class=ctd><%=icount%></td>
  <td class=ctd alt="<%=replace(Rs("rwlr"),vbcrlf,"<br>")%>"><%=Rs("id")%></td>
  <td class=ctd><%=Rs("rwlx")%></td>
  <td class=ctd><%=Rs("zrr")%></td>
  <td class=ctd><%=xjDate(Rs("jssj"),1)%></td>
  <td class=ctd><%=Rs("zf")%></td>
  <td class=ctd>&nbsp;</td>
  <td class=ctd>&nbsp;</td>
</tr>
<%
			icount = icount + 1
			ilxrwzf=ilxrwzf+Rs("zf")
			Rs.movenext
		loop
		%>
<tr>
  <td class=rtd colspan=5>���������ܷ�:</td>
  <td class=ctd colspan=3><b><%=ilxrwzf%></b></td>
</tr>
<%
	end If
	Rs.close
end function

function total_mt()		'�ܷ�ͳ��
	Dim iTotalFz
	If Fix(ilxrwzf + irwzf + iaddfz)<(ilxrwzf + irwzf + iaddfz) Then
		iTotalFz=Fix(ilxrwzf + irwzf + iaddfz) + 1
	Else
		iTotalFz=Fix(ilxrwzf + irwzf + iaddfz)
	End If
%>
<tr>
  <td class=rtd colspan=5>�ܷ�:</td>
  <td class=ctd colspan=3><b><%=iTotalFz%></b></td>
</tr>
<%
end function
%>
<script language="javascript">
function changesjf(arg){
var strsjf
strsjf=showModalDialog("mtstat_c.asp?id="+arg, "", "dialogWidth:280px; dialogHeight:160px; center:yes; help: no; scroll: no; status:no;");
if (strsjf==2){
window.location.reload();
}
}
</script>
