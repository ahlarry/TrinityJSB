<!--#include file="include/conn.asp"-->
<%
'17:18 2006-12-1-������
Call ChkPageAble(0)
Call ChkDepart("������")
CurPage="ģ�ߵ��� �� �鿴������Ϣ"					'ҳ�������λ��( ��������� �� ���������)
strPage="mtest"
'Call FileInc(0, "js/mtest.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchLsh()%></td>
  </tr>
  <Tr>
    <Td class=ctd height=300><%Call mtestDisplay()%></Td>
  </Tr>
</Table>
<%
End Sub

Function mtestDisplay()
	Dim s_lsh
	s_lsh=UCase(Trim(Request("s_lsh")))
	If s_lsh="" Then Call TbTopic("������鿴������Ϣģ�ߵ���ˮ��!") : Exit Function
	strSql="select a.*, b.*,a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh='"&s_lsh&"' and a.lsh=b.lsh"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.eof or Rs.bof Then
		Call JsAlert("��ˮ�� ��" & s_lsh & "�� �����鲻���ڻ�����ֲ�û�����! ","mtest_display.asp")
	Else
		Call mould_inf(Rs)
		Response.write(XjLine(10, "100%", ""))
		Call mtest_display(s_lsh)
		Response.write(XjLine(10, "100%", ""))
		Call PreNext(s_lsh)
		Response.write(XjLine(10, "100%", ""))
	End If
	Rs.close
End Function

Function mould_inf(Rs)
%>
<%Call TbTopic("��ˮ�� "&Rs("lsh")&" ģ����Ϣ")%>
<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
  <tr>
    <td class=th width="10%">������</td>
    <td class=th width="*">��������</td>
    <td class=th width="10%">��ˮ��</td>
    <td class=th width="10%">��λ����</td>
    <td class=th width="10%">��������</td>
    <td class=th width="15%">���Կ�ʼ</td>
    <td class=th width="15%">�������</td>
    <td class=th width="10%">���Դ���</td>
  </tr>
  <tr>
    <td class=ctd><%=Rs("ddh")%></td>
    <td class=ctd><%=Rs("dmmc")%></td>
    <td class=ctd><a href="mtask_display.asp?s_lsh=<%=Rs("lsh")%>"><%=Rs("lsh")%></a></td>
    <td class=ctd><%=Rs("dwmc")%></td>
    <td class=ctd><%=Rs("mjxx")%></td>
    <td class=ctd><%=xjDate(Rs("tskssj"),1)%>&nbsp;</td>
    <td class=ctd alt="<%If isnull(Rs("tsjssj")) Then%>���ڵ���<%Else%>���Խ���<%End If%>"><%=xjDate(Rs("tsgxsj"),1)%>&nbsp;</td>
    <td class=ctd><%=xjweb.RsCount("ts_tsxx where lsh='"&Rs("lsh")&"' and not(ps)")%></td>
  </tr>
</table>
<%
End Function

Function mtest_display(lsh)
	Dim prs, itscs, ipscs
	strSql="select * from [ts_tsxx] where lsh='"&lsh&"' order by id desc"
	itscs=xjweb.rscount("[ts_tsxx] where lsh='"&lsh&"' and not(ps)")
	ipscs=xjweb.rscount("[ts_tsxx] where lsh='"&lsh&"' and ps")
	Set prs = xjweb.Exec(strSql, 1)
	If Prs.Eof Or Prs.Bof Then Prs.Close : Set Prs=Nothing : Call TbTopic("��ʱû���κε�����Ϣ!") : Exit Function
	Call TbTopic("��ˮ�� " &lsh&" ģ�ߵ�����Ϣ�б�")
%>
<table class=xtable cellspacing=0 cellpadding=3 width="95%" align="center">
  <%
	do while not prs.eof
		If prs("ps") Then
	%>
  <tr bgcolor=#dddddd>
    <td class=ctd width="10%" rowspan="3">�� <b><%=ipscs%></b> ��<br>
      ����</td>
    <td class=rtd width="10%">��������:</td>
    <td class=ltd width="*"><%=xjweb.htmltocode(prs("tslr"))%></td>
  </tr>
  <tr bgcolor=#dddddd>
    <td class=rtd>������:</td>
    <td class=ltd><%=xjweb.htmltocode(prs("tsyy"))%></td>
  </tr>
  <form action="mtest_indb.asp?action=delete" method=post onsubmit="return confirm('ȷ��ɾ����?');">
    <tr bgcolor=#dddddd>
      <td class=rtd colspan="2">ǩд:<%=prs("tsr")%> ����:<%=prs("tssj")%>
        <%If chkable(6) and prs("tsr")=Session("userName") Then%>
        &nbsp;<a href="mtest_change.asp?id=<%=prs("id")%>&cs=<%=ipscs%>&s_lsh=<%=lsh%>&ps=true">�༩</a>&nbsp;
        <%End If%>
        <%If chkable(1) Then%>
        <input type="submit" value=" ɾ�� ">
        <input type="hidden" name=id value="<%=prs("id")%>">
        <input type="hidden" name="lsh" value="<%=prs("lsh")%>">
        <%End If%></td>
    </tr>
  </form>
  <%
			ipscs=ipscs-1
		Else
	%>
  <tr>
    <td class=ctd width="10%" rowspan="3">�� <b><%=itscs%></b> ��</td>
    <td class=rtd width="10%">����ԭ��:</td>
    <td class=ltd width="*"><%=xjweb.htmltocode(prs("tsyy"))%></td>
  </tr>
  <tr>
    <td class=rtd>��������:</td>
    <td class=ltd><%=xjweb.htmltocode(prs("tslr"))%></td>
  </tr>
  <form action="mtest_indb.asp?action=delete" method=post onsubmit="return confirm('ȷ��ɾ����?');">
    <tr>
      <td class=rtd colspan="2">����:<%=prs("tsr")%> ����:<%=prs("tssj")%>
        <%If chkable(6) and prs("tsr")=Session("userName") Then%>
        &nbsp;<a href="mtest_change.asp?id=<%=prs("id")%>&cs=<%=itscs%>&s_lsh=<%=lsh%>&ps=false">�༩</a>&nbsp;
        <%End If%>
        <%If chkable(1) Then%>
        <input type="submit" value=" ɾ�� ">
        <input type="hidden" name="id" value="<%=prs("id")%>">
        <input type="hidden" name="lsh" value="<%=prs("lsh")%>">
        <%End If%></td>
    </tr>
  </form>
  <%
			itscs=itscs-1
		End If
		prs.movenext
	loop
	prs.close
	Set prs = nothing
	%>
</table>
<%
End Function

Function PreNext(ilsh)
Dim strOrder,strPre,strNext,TmpSql,Trs,Tsj
strOrder=Trim(Request("order")) : strPre="" : strNext="" : Tsj=""

	TmpSql="select * from [ts_mould] where  lsh = '" &ilsh& "'"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	Tsj=Trs("tsgxsj")
	Trs.close
	Set Trs = nothing

If strOrder="tsgxsj" Then
	TmpSql="select * from [ts_mould] where datediff('s',tsgxsj,'"&Tsj&"')>0 order by tsgxsj desc,lsh desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then
		strPre="Beg"
	Else
		strPre=Trs("lsh")
	End  If
	TmpSql="select * from [ts_mould] where datediff('s',tsgxsj,'"&Tsj&"')<0 order by tsgxsj,lsh desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then
		strNext="End"
	Else
		strNext=Trs("lsh")
	End  If
Else
	TmpSql="select a.*, b.*,a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh < '"&ilsh&"' and a.lsh=b.lsh order by a.lsh desc,tsgxsj desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then
		strPre="Beg"
	Else
		strPre=Trs("lsh")
	End  If
	TmpSql="select a.*, b.*,a.lsh as lsh from [mtask] a, [ts_mould] b where a.lsh > '"&ilsh&"' and a.lsh=b.lsh order by a.lsh,tsgxsj desc"
	Set Trs = Server.Createobject("adodb.Recordset")
	Trs.Open TmpSql,Conn,1,3
	If Trs.BOF Then
		strNext="End"
	Else
		strNext=Trs("lsh")
	End  If
End If
Trs.close
Set Trs = nothing
%>
<table cellspacing=0 cellpadding=3 width="95%" align="center">
  <tr>
    <td width="20%">
    <%If strPre="Beg" Then
    	Response.write("")
    else%>
   		<a href=mtest_display.asp?s_lsh=<%=strPre%>&order=<%=strOrder%>><strong>��һ����<%=strPre%></strong></a>
   	<%End If%>
    </td>
    <td width="*" align="center">����:
      <select name="order" onchange='location.href("<%=Request.servervariables("script_name")%>?s_lsh=<%=ilsh%>&order=" + this.value);'>
        <option value="" selected="selected">�� ˮ ��</option>
        <option value="tsgxsj" <%If strOrder="tsgxsj" Then%>selected<%End If%>>�������</option>
      </select>��ע����ˮ���е�"<font size="4" color="#ff0000"><strong>C</strong></font>"</td>
    <td width="20%" align="right">
    <%If strNext="End" Then
    	Response.write("")
    else%>
    	<a href=mtest_display.asp?s_lsh=<%=strNext%>&order=<%=strOrder%>><strong>��һ����<%=strNext%></strong></a>
    <%End If%>
    </td>
  </tr>
</Table>
<%End Function%>
