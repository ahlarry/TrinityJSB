<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'16:10 2007-1-26-������
Call ChkPageAble("3,4")
CurPage="������� �� ����������"
strPage="mtask"
Call FileInc(0, "js/mtask.js")
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, sjcount, dbcount
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder
sjcount=0
dbcount=0

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<table class="xtable" cellspacing="0" cellpadding="0" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd" style="border-right-style:none"><%Call SearchLsh()%></td>
    <td class="rtd" style="border-left-style:none"><%Call Taxis()%></td>
  </tr>
  <tr>
    <td class="ctd" height="300" colspan="2"><%Call mtaskAssign()%>
      <%Response.Write(XjLine(5,"100%",""))%>
      <%Response.Write(XjLine(1,"100%","#00659c"))%>
      <%Response.Write(XjLine(5,"100%",""))%>
      <%Call task_assign_nofinished()%>
      <%Call jsdb_nofinished()%>
<table width="99%" cellpadding="2" cellspacing="0" border="0">
  <tr>
    <td class="td_ltd">�����������������&nbsp;<%=sjcount%>&nbsp;��,��������������&nbsp;<%=dbcount%>&nbsp;��</td>
  </tr>
</table>
    </td>
  </tr>
</table>
<%
End Sub

Function Taxis()
Dim strO
strO="1"
%>
����:
<select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;order=&quot; + this.value);'>
  <option value="jhjssj" selected="selected">�ƻ��������ʱ��</option>
  <option value="ddh" <%If strOrder="ddh" Then%>selected<%End If%>>������</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>��ˮ��</option>
</select>
<%
End Function
Function mtaskAssign()
	Dim s_lsh,s_hth
	s_lsh="" : s_hth=""
	If Trim(request("s_lsh"))<>"" Then s_lsh=Trim(request("s_lsh"))
	If Trim(request("s_hth"))<>"" Then s_hth=Trim(request("s_hth"))
	If s_lsh="" and s_hth="" Then Call TbTopic("�������ѡ����������������ˮ��!") : Exit Function
	if s_lsh<>"" Then
		strSql="select * from [mtask] where lsh='"&s_lsh&"'"
		Set Rs=xjweb.Exec(strSql,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("��ˮ�� " & s_lsh & " �����鲻����! ����������!","mtask_assign.asp")
		Else
			If Not isnull(rs("sjjssj")) then
				Call JsAlert("�������Ѿ����,����Ҫ�ٽ��з���!","mtask_assign.asp")
				'17:18 2007-3-20-���ڶ�
			elseif rs("zz")=session("userName") or rs("jgzz")=session("userName") or rs("sjzz")=session("userName") or rs("group")=session("userGroup") or chkable(3) then
				call mtask_assign(rs)
			elseif rs("zz")<>"" Then
				Call JsAlert("��������鳤�� " & rs("zz") &" ! ����ϵ "&rs("zz")&" �����������!","")
				else
					Call JsAlert("��������鳤�� " & rs("jgzz") &"(�ṹ)��"& rs("sjzz") &"(���) ! ����ϵ�鳤�����������!","")
			end if
		end if
	else
		strSql="select * from [jsdb] where hth='"&s_hth&"'"
		Set Rs=xjweb.Exec(strSql,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("��ͬ�� " & s_hth & " �ļ������������鲻����! ����������!","mtask_assign.asp")
		Else
			If Not isnull(rs("shjssj")) then
				Call JsAlert("�������Ѿ����,����Ҫ�ٽ��з���!","mtask_assign.asp")
			elseif rs("zz")=session("userName") or chkable(3) then
					call jsdb_assign(rs)
				else
					Call JsAlert("��������鳤�� " & rs("zz") &" ! ����ϵ "&rs("zz")&" �����������!","")
			end if
		end if
	end if
	rs.close
end function

function mtask_assign(rs)
	call mtask_fewinfo(rs)
	Response.Write(XjLine(4,"100%",""))
	call mtask_userinfo(rs)
	Response.Write(XjLine(4,"100%",""))
	Response.Write(XjLine(1,"100%","#00659c"))
	Response.Write(XjLine(4,"100%",""))

	if not isnull(rs("fsjs")) then
		If chkable(3) Then
			call mtask_finish(rs)
		else
			Call JsAlert("�������Ѹ������! �뼰ʱ�����Ӧ���Ե�����!","mtask_assign.asp")
		End If
	else
		select case rs("mjxx")
			case "ȫ��"
				if not isnull(rs("mtbomjs")) and not isnull(rs("dxbomjs")) then
					call mtask_audit(rs)
				else
					If IsNull(Rs("dxshr")) and IsNull(Rs("mtshr")) and (rs("rwlr")="���") Then
						call mtask_nassign(rs)
					else
						call mtask_assigntask(rs)
					End If
				end if
			case "ģͷ"
				if not isnull(rs("mtbomjs")) then
					call mtask_audit(rs)
				else
					If IsNull(Rs("mtshr")) and (rs("rwlr")="���") Then
						call mtask_nassign(rs)
					else
						call mtask_assigntask(rs)
					End If
				end if
			case "����"
				if not isnull(rs("dxbomjs")) then
					call mtask_audit(rs)
				else
					If IsNull(Rs("dxshr")) and (rs("rwlr")="���") Then
						call mtask_nassign(rs)
					else
						call mtask_assigntask(rs)
					End If
				end if
		end select
	end if
end function		'mtask_assign()

Function mtask_finish(rs)
	if (rs("mjxx")="ȫ��" and  (isnull(rs("mttsdjs")) or isnull(rs("dxtsdjs")))) or  (rs("mjxx")="ģͷ" and isnull(rs("dxtsdjs"))) or  (rs("mjxx")="����" and isnull(rs("dxtsdjs"))) then
		Call JsAlert("�������鳤��ʱ�����Ӧ���Ե�����!","mtask_assign.asp")
	end if
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%" align="center">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="if(document.all.zrpsjl.value==''){alert('����д�����¼!');return false;}">
    <tr>
      <td class="ctd" colspan="2" >ģ�߷�ֵ:<b><%=Rs("mjzf")%></b></td>
    </tr>
    <tr>
      <td class="ctd" width="15%">�����¼:</td>
      <td class="ctd"><textarea name="zrpsjl" rows="7" cols="60"><%=Rs("psjl")%></textarea></td>
    </tr>
    <tr>
      <td class="ctd">����ʱ��</td>
      <td class="ctd"><select id="psd" name="psd">
          <%for i = DateAdd("m", -2, date()) to date()%>
          <%if i = date() then%>
          <option value='<%=i%>' selected="selected"><%=i%></option>
          <%else%>
          <option value='<%=i%>'><%=i%></option>
          <%end if%>
          <%next%>
        </select>
        <input name="fplr" type="submit" value="ȫ�׽���" />
      </td>
    </tr>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<%
End Function

function mtask_audit(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%" align="center">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="if(document.all.psjl.value==''){alert('����д�����¼!');return false;}">
    <tr>
      <td class="ctd"  width="15%">�����¼:</td>
      <td class="ctd"><textarea name="psjl" rows="5" cols="80"><%If datediff("d", now, rs("jhjssj")) > 3 Then%>ΪԼ�������ź�����,ϵͳ��������ǰ�ƻ�3���������.<%End If%>
</textarea></td>
      <td class="ctd"><input name="fplr" type="submit" value="��������" <%If datediff("d", now, rs("jhjssj")) > 3 Then%> disabled="disabled" <%End If%> /></td>
    </tr>
    <tr>
    	<td class="ctd"  width="15%">������:</td>
    	<td class="ctd">
         <select name="zrr">
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>' <%if session("userName")=c_allzz(i) then %> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
          <option value=�����ڡ�>����</option>
          <option value=����Сͣ��>��Сͣ</option>
        </select>
	</td>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<%
end function
'����������������������������������������������������������������������������������������������
function mtask_nassign(rs)
%>
<table border="0" cellpadding="3" cellspacing="0">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style="font-size:14px;font-weight:bold;">������ˮ�� <%=rs("lsh")%> ������:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
	select case rs("mjxx") & rs("rwlr")
		case "ȫ�����"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'���ṹ�鳤�ܷ���ṹ����
			if Rs("HGJ")>0 Then
				if isnull(rs("mtjgr")) and isnull(rs("dxjgr")) and isnull(rs("gjjgr")) then sel_opt("��ʼȫ�׽ṹ")
				elseif isnull(rs("mtjgr")) and isnull(rs("dxjgr")) then sel_opt("��ʼȫ�׽ṹ")
			End if
			if isnull(rs("mtjgr")) then sel_opt("��ʼģͷ�ṹ")
			if isnull(rs("dxjgr")) then sel_opt("��ʼ���ͽṹ")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("��ʼ�󹲼��ṹ")
			if Rs("HGJ")>0 Then
				if rs("mtjgr")=rs("dxjgr") and rs("dxjgr")=rs("gjjgr") and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) and isnull(rs("gjjgjs")) then sel_opt("����ȫ�׽ṹ")
				elseif (rs("mtjgr")=rs("dxjgr")) and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) then sel_opt("����ȫ�׽ṹ")
			End if
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("����ģͷ�ṹ")
			if (not isnull(rs("dxjgr"))) and isnull(rs("dxjgjs")) then sel_opt("�������ͽṹ")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("�����󹲼��ṹ")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjjgjs"))) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) and isnull(rs("mtjgshr")) and isnull(rs("dxjgshr")) and isnull(rs("gjjgshr")) then sel_opt("��ʼȫ�׽ṹȷ��")
				elseif isnull(rs("mtjgshr")) and isnull(rs("dxjgshr")) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) then sel_opt("��ʼȫ�׽ṹȷ��")
			End if
			if isnull(rs("mtjgshr")) and not isnull(rs("mtjgjs")) then sel_opt("��ʼģͷ�ṹȷ��")
			if isnull(rs("dxjgshr")) and not isnull(rs("dxjgjs")) then sel_opt("��ʼ���ͽṹȷ��")
			if isnull(rs("gjjgshr")) and not isnull(rs("gjjgjs")) then sel_opt("��ʼ�󹲼��ṹȷ��")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'������鳤�ܷ����������
			if Rs("HGJ")>0 Then
				if rs("mtjgshr")=rs("dxjgshr") and rs("dxjgshr")=rs("gjjgshr") and isnull(rs("mtjgshjs")) and isnull(rs("dxjgshjs")) and isnull(rs("gjjgshjs")) then sel_opt("����ȫ�׽ṹȷ��")
				elseif (rs("mtjgshr")=rs("dxjgshr")) and isnull(rs("mtjgshjs")) and isnull(rs("dxjgshjs")) then sel_opt("����ȫ�׽ṹȷ��")
			End if
			if not isnull(rs("mtjgshr")) and isnull(rs("mtjgshjs")) then sel_opt("����ģͷ�ṹȷ��")
			if not isnull(rs("dxjgshr")) and isnull(rs("dxjgshjs")) then sel_opt("�������ͽṹȷ��")
			if not isnull(rs("gjjgshr")) and isnull(rs("gjjgshjs")) then sel_opt("�����󹲼��ṹȷ��")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjjgshjs"))) and (not isnull(rs("mtjgshjs"))) and (not isnull(rs("dxjgshjs"))) and isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and isnull(rs("gjsjr")) then sel_opt("��ʼȫ�����")
				elseif isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and (not isnull(rs("mtjgshjs"))) and (not isnull(rs("dxjgshjs"))) then sel_opt("��ʼȫ�����")
			End if
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgshjs")) then sel_opt("��ʼģͷ���")
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgshjs")) then sel_opt("��ʼ�������")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgshjs")) then sel_opt("��ʼ�󹲼����")

			if Rs("HGJ")>0 Then
				if rs("mtsjr")=rs("dxsjr") and rs("dxsjr")=rs("gjsjr") and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) and isnull(rs("gjsjjs")) then sel_opt("����ȫ�����")
				elseif (rs("mtsjr")=rs("dxsjr")) and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) then sel_opt("����ȫ�����")
			End if
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("����ģͷ���")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("�����������")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("�����󹲼����")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjsjjs"))) and (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtsjshr")) and isnull(rs("dxsjshr")) and isnull(rs("gjsjshr")) then sel_opt("��ʼȫ�����ȷ��")
				elseif (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtsjshr")) and isnull(rs("dxsjshr")) then sel_opt("��ʼȫ�����ȷ��")
			End if
			if isnull(rs("mtsjshr")) and not isnull(rs("mtsjjs")) then sel_opt("��ʼģͷ���ȷ��")
			if isnull(rs("dxsjshr")) and not isnull(rs("dxsjjs")) then sel_opt("��ʼ�������ȷ��")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjsjshr")) then sel_opt("��ʼ�󹲼����ȷ��")

			if Rs("HGJ")>0 Then
				if rs("mtsjshr")=rs("dxsjshr") and rs("dxsjshr")=rs("gjsjshr") and isnull(rs("mtsjshjs")) and isnull(rs("dxsjshjs")) and isnull(rs("gjsjshjs")) then sel_opt("����ȫ�����ȷ��")
				elseif (rs("mtsjshr")=rs("dxsjshr")) and isnull(rs("mtsjshjs")) and isnull(rs("dxsjshjs")) then sel_opt("����ȫ�����ȷ��")
			End if
			if not isnull(rs("mtsjshr")) and isnull(rs("mtsjshjs")) then sel_opt("����ģͷ���ȷ��")
			if not isnull(rs("dxsjshr")) and isnull(rs("dxsjshjs")) then sel_opt("�����������ȷ��")
			if (not isnull(rs("gjsjshr"))) and isnull(rs("gjsjshjs")) then sel_opt("�����󹲼����ȷ��")
		End If

'       	if not(isnull(rs("mtsjshjs"))) and not(isnull(rs("dxsjshjs"))) and not(isnull(rs("gjsjshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("ȫ�׹������")
'       	if not(isnull(rs("mtsjshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'       	if not(isnull(rs("dxsjshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'       	if not(isnull(rs("gjsjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("�����������")
'       	if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr")))  and not(isnull(rs("gjgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("ȫ�׹������")
'       	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")
'       	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")
'       	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("�����������")
		if isnull(rs("mtbomr")) and not isnull(rs("mtsjshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxsjshjs")) then sel_opt("��ʼȫ��BOM")
		if isnull(rs("mtbomr")) and not isnull(rs("mtsjshjs")) then sel_opt("��ʼģͷBOM")
		if isnull(rs("dxbomr")) and not isnull(rs("dxsjshjs")) then sel_opt("��ʼ����BOM")

		if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("����ȫ��BOM")
		if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")
		if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")

		case "ģͷ���"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'���ṹ�鳤�ܷ���ṹ����
			if isnull(rs("mtjgr")) then sel_opt("��ʼģͷ�ṹ")
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("����ģͷ�ṹ")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("��ʼ�󹲼��ṹ")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("�����󹲼��ṹ")
			if isnull(rs("mtjgshr")) and not isnull(rs("mtjgjs")) then sel_opt("��ʼģͷ�ṹȷ��")
			if isnull(rs("gjjgshr")) and not isnull(rs("gjjgjs")) then sel_opt("��ʼ�󹲼��ṹȷ��")
			if not isnull(rs("gjjgshr")) and isnull(rs("gjjgshjs")) then sel_opt("�����󹲼��ṹȷ��")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'������鳤�ܷ����������
			if (not isnull(rs("mtjgshr"))) and isnull(rs("mtjgshjs")) then sel_opt("����ģͷ�ṹȷ��")
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgshjs")) then sel_opt("��ʼģͷ���")
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("����ģͷ���")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgshjs")) then sel_opt("��ʼ�󹲼����")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("�����󹲼����")
			if isnull(rs("mtsjshr")) and not isnull(rs("mtsjjs")) then sel_opt("��ʼģͷ���ȷ��")
			if not isnull(rs("mtsjshr")) and isnull(rs("mtsjshjs")) then sel_opt("����ģͷ���ȷ��")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjsjshr")) then sel_opt("��ʼ�󹲼����ȷ��")
			if (not isnull(rs("gjsjshr"))) and isnull(rs("gjsjshjs")) then sel_opt("�����󹲼����ȷ��")
		End If

'          	if not(isnull(rs("mtsjshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'          	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")
			if isnull(rs("mtbomr")) and not isnull(rs("mtsjshjs")) then sel_opt("��ʼģͷBOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")

		case "�������"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'���ṹ�鳤�ܷ���ṹ����
			if isnull(rs("dxjgr")) then sel_opt("��ʼ���ͽṹ")
			if not isnull(rs("dxjgr")) and isnull(rs("dxjgjs")) then sel_opt("�������ͽṹ")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("��ʼ�󹲼��ṹ")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("�����󹲼��ṹ")
			if isnull(rs("dxjgshr")) and not isnull(rs("dxjgjs")) then sel_opt("��ʼ���ͽṹȷ��")
			if isnull(rs("gjjgshr")) and not isnull(rs("gjjgjs")) then sel_opt("��ʼ�󹲼��ṹȷ��")
			if not isnull(rs("gjjgshr")) and isnull(rs("gjjgshjs")) then sel_opt("�����󹲼��ṹȷ��")
		End If


		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'������鳤�ܷ����������
			if not isnull(rs("dxjgshr")) and isnull(rs("dxjgshjs")) then sel_opt("�������ͽṹȷ��")
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgshjs")) then sel_opt("��ʼ�������")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("�����������")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgshjs")) then sel_opt("��ʼ�󹲼����")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("�����󹲼����")
			if isnull(rs("dxsjshr")) and not isnull(rs("dxsjjs")) then sel_opt("��ʼ�������ȷ��")
			if not isnull(rs("dxsjshr")) and isnull(rs("dxsjshjs")) then sel_opt("�����������ȷ��")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjsjshr")) then sel_opt("��ʼ�󹲼����ȷ��")
			if (not isnull(rs("gjsjshr"))) and isnull(rs("gjsjshjs")) then sel_opt("�����󹲼����ȷ��")
		End If

'          	if not(isnull(rs("dxsjshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'          	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")
			if isnull(rs("dxbomr")) and not isnull(rs("dxsjshjs")) then sel_opt("��ʼ����BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")

		case else
			response.write(rs("mjxx") & rs("rwlr"))
	end select
%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type="submit" value=" ���� " /></td>
    </tr>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<script language="JavaScript" type="text/javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='����')
				objdoc.form_assign.zrr.disabled=true;
			else
				objdoc.form_assign.zrr.disabled=false;
		}
		checkjs();

		function checksubinf(frm)
		{
			if (frm.fplr.value==""){alert("��ѡ���������!"); frm.fplr.focus(); return false;}
			if ((frm.zrr.value=="") && (!frm.zrr.disabled)){alert("��ѡ��������!"); frm.zrr.focus(); return false;}
			return true;
		}
	</script>
<%
end function
'����������������������������������������������������������������������������������������������
function mtask_assigntask(rs)
%>
<table border="0" cellpadding="3" cellspacing="0">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style="font-size:14px;font-weight:bold;">������ˮ�� <%=rs("lsh")%> ������:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
	select case rs("mjxx") & rs("rwlr")
		case "ȫ�����"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'���ṹ�鳤�ܷ���ṹ����
			if Rs("HGJ")>0 Then
				if isnull(rs("mtjgr")) and isnull(rs("dxjgr")) and isnull(rs("gjjgr")) then sel_opt("��ʼȫ�׽ṹ")
				elseif isnull(rs("mtjgr")) and isnull(rs("dxjgr")) then sel_opt("��ʼȫ�׽ṹ")
			End if
			if isnull(rs("mtjgr")) then sel_opt("��ʼģͷ�ṹ")
			if isnull(rs("dxjgr")) then sel_opt("��ʼ���ͽṹ")
			if Rs("HGJ")>0 and isnull(rs("gjjgr")) then sel_opt("��ʼ�󹲼��ṹ")

			if Rs("HGJ")>0 Then
				if rs("mtjgr")=rs("dxjgr") and rs("dxjgr")=rs("gjjgr") and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) and isnull(rs("gjjgjs")) then sel_opt("����ȫ�׽ṹ")
				elseif (rs("mtjgr")=rs("dxjgr")) and isnull(rs("mtjgjs")) and isnull(rs("dxjgjs")) then sel_opt("����ȫ�׽ṹ")
			End if
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("����ģͷ�ṹ")
			if (not isnull(rs("dxjgr"))) and isnull(rs("dxjgjs")) then sel_opt("�������ͽṹ")
			if (not isnull(rs("gjjgr"))) and isnull(rs("gjjgjs")) then sel_opt("�����󹲼��ṹ")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'������鳤�ܷ����������
			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjjgjs"))) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) and isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and isnull(rs("gjsjr")) then sel_opt("��ʼȫ�����")
				elseif isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and (not isnull(rs("mtjgjs"))) and (not isnull(rs("dxjgjs"))) then sel_opt("��ʼȫ�����")
			End if
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgjs")) then sel_opt("��ʼģͷ���")
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgjs")) then sel_opt("��ʼ�������")
			if isnull(rs("gjsjr")) and not isnull(rs("gjjgjs")) then sel_opt("��ʼ�󹲼����")

			if Rs("HGJ")>0 Then
				if rs("mtsjr")=rs("dxsjr") and rs("dxsjr")=rs("gjsjr") and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) and isnull(rs("gjsjjs")) then sel_opt("����ȫ�����")
				elseif (rs("mtsjr")=rs("dxsjr")) and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) then sel_opt("����ȫ�����")
			End if
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("����ģͷ���")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("�����������")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("�����󹲼����")
		End If

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjsjjs"))) and (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) and isnull(rs("gjshr")) then sel_opt("��ʼȫ�����")
				elseif (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) then sel_opt("��ʼȫ�����")
			End if
			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("��ʼģͷ���")
			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("��ʼ�������")
			if  (not isnull(rs("gjsjjs"))) and isnull(rs("gjshr")) then sel_opt("��ʼ�󹲼����")

			if Rs("HGJ")>0 Then
				if rs("mtshr")=rs("dxshr") and rs("dxshr")=rs("gjshr") and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) and isnull(rs("gjshjs")) then sel_opt("����ȫ�����")
				elseif (rs("mtshr")=rs("dxshr")) and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) then sel_opt("����ȫ�����")
			End if
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("����ģͷ���")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("�����������")
			if (not isnull(rs("gjshr"))) and isnull(rs("gjshjs")) then sel_opt("�����󹲼����")

'     	  	if not(isnull(rs("mtshjs"))) and not(isnull(rs("dxshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("ȫ�׹������")
'   	    	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'      	 	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'       		if not(isnull(rs("gjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("�����������")
'      	 	if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("ȫ�׹������")
'      	 	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")
'     	  	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")
'     	  	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("�����������")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼȫ��BOM")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("��ʼģͷBOM")
			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼ����BOM")
			if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("����ȫ��BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")

		case "ȫ�׸���"
			if Rs("HGJ")>0 Then
				if isnull(rs("mtsjr")) and isnull(rs("dxsjr")) and isnull(rs("gjsjr")) then sel_opt("��ʼȫ�׸���")
				elseif isnull(rs("mtsjr")) and isnull(rs("dxsjr")) then sel_opt("��ʼȫ�׸���")
			End if
			if isnull(rs("mtsjr")) then sel_opt("��ʼģͷ����")
			if isnull(rs("dxsjr")) then sel_opt("��ʼ���͸���")
			if Rs("HGJ")>0 and isnull(rs("gjsjr")) then sel_opt("��ʼ�󹲼�����")

			if Rs("HGJ")>0 Then
				if rs("mtsjr")=rs("dxsjr") and rs("dxsjr")=rs("gjsjr") and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) and isnull(rs("gjsjjs")) then sel_opt("����ȫ�׸���")
				elseif (rs("mtsjr")=rs("dxsjr")) and isnull(rs("mtsjjs")) and isnull(rs("dxsjjs")) then sel_opt("����ȫ�׸���")
			End if
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("����ģͷ����")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("�������͸���")
			if not isnull(rs("gjsjr")) and isnull(rs("gjsjjs")) then sel_opt("�����󹲼�����")

			if Rs("HGJ")>0 Then
				if (not isnull(rs("gjsjjs"))) and (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) and isnull(rs("gjshr")) then sel_opt("��ʼȫ�����")
				elseif (not isnull(rs("mtsjjs"))) and (not isnull(rs("dxsjjs"))) and isnull(rs("mtshr")) and isnull(rs("dxshr")) then sel_opt("��ʼȫ�����")
			End if
			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("��ʼģͷ���")
			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("��ʼ�������")
			if isnull(rs("gjshr")) and not isnull(rs("gjsjjs")) then sel_opt("��ʼ�󹲼����")

			if Rs("HGJ")>0 Then
				if rs("mtshr")=rs("dxshr") and rs("dxshr")=rs("gjshr") and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) and isnull(rs("gjshjs")) then sel_opt("����ȫ�����")
				elseif (rs("mtshr")=rs("dxshr")) and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) then sel_opt("����ȫ�����")
			End if
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("����ģͷ���")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("�����������")
			if not isnull(rs("gjshr")) and isnull(rs("gjshjs")) then sel_opt("�����󹲼����")

'  	     	if not(isnull(rs("mtshjs"))) and not(isnull(rs("dxshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("ȫ�׹������")
'    	   	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'    	   	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'    	   	if not(isnull(rs("gjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("�����������")
'   	    	if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("ȫ�׹������")
'   	    	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")
'   	    	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")
'   	    	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("�����������")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼȫ��BOM")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("��ʼģͷBOM")
			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼ����BOM")

			if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("����ȫ��BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")

		case "ȫ�׸���"
			if Rs("HGJ")>0 Then
				if isnull(rs("mtshr")) and isnull(rs("dxshr")) and isnull(rs("gjshr")) then sel_opt("��ʼȫ�׸���")
				elseif isnull(rs("mtshr")) and isnull(rs("dxshr")) then sel_opt("��ʼȫ�׸���")
			End if
			if isnull(rs("mtshr")) then sel_opt("��ʼģͷ����")
			if isnull(rs("dxshr")) then sel_opt("��ʼ���͸���")
			if Rs("HGJ")>0 and isnull(rs("gjshr")) then sel_opt("��ʼ�󹲼�����")

			if Rs("HGJ")>0 Then
				if rs("mtshr")=rs("dxshr") and rs("dxshr")=rs("gjshr") and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) and isnull(rs("gjshjs")) then sel_opt("����ȫ�׸���")
				elseif (rs("mtshr")=rs("dxshr")) and isnull(rs("mtshjs")) and isnull(rs("dxshjs")) then sel_opt("����ȫ�׸���")
			End if
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("����ģͷ����")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("�������͸���")
			if not isnull(rs("gjshr")) and isnull(rs("gjshjs")) then sel_opt("�����󹲼�����")

'    	   	if not(isnull(rs("mtshjs"))) and not(isnull(rs("dxshjs"))) and isnull(rs("mtgysjr")) and isnull(rs("dxgysjr")) and isnull(rs("gjgysjr")) then sel_opt("ȫ�׹������")
'       		if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'       		if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'       		if not(isnull(rs("gjshjs"))) and isnull(rs("gjgysjr")) then sel_opt("�����������")
'       		if not(isnull(rs("mtgysjr"))) and not(isnull(rs("dxgysjr"))) and isnull(rs("mtgyshr")) and isnull(rs("dxgyshr")) and isnull(rs("gjgyshr")) then sel_opt("ȫ�׹������")
'       		if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")
'      	 	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")
'      	 	if not(isnull(rs("gjgysjr"))) and isnull(rs("gjgyshr")) then sel_opt("�����������")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) and isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼȫ��BOM")
			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("��ʼģͷBOM")
			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼ����BOM")

			if (rs("mtbomr")=rs("dxbomr")) and isnull(rs("mtbomjs")) and isnull(rs("dxbomjs")) then sel_opt("����ȫ��BOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")

		case "ģͷ���"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'���ṹ�鳤�ܷ���ṹ����
			if isnull(rs("mtjgr")) then sel_opt("��ʼģͷ�ṹ")
			if (not isnull(rs("mtjgr"))) and isnull(rs("mtjgjs")) then sel_opt("����ģͷ�ṹ")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'������鳤�ܷ����������
			if isnull(rs("mtsjr")) and not isnull(rs("mtjgjs")) then sel_opt("��ʼģͷ���")
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("����ģͷ���")
		End If

			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("��ʼģͷ���")
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("����ģͷ���")

'      	 	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'       	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("��ʼģͷBOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")

		case "ģͷ����"
			if isnull(rs("mtsjr")) then sel_opt("��ʼģͷ����")
			if not isnull(rs("mtsjr")) and isnull(rs("mtsjjs")) then sel_opt("����ģͷ����")

			if isnull(rs("mtshr")) and not isnull(rs("mtsjjs")) then sel_opt("��ʼģͷ���")
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("����ģͷ���")

' 	      	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'   	    	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("��ʼģͷBOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")

		case "ģͷ����"
			if isnull(rs("mtshr")) then sel_opt("��ʼģͷ����")
			if not isnull(rs("mtshr")) and isnull(rs("mtshjs")) then sel_opt("����ģͷ����")

'    	   	if not(isnull(rs("mtshjs"))) and isnull(rs("mtgysjr")) then sel_opt("ģͷ�������")
'   	    	if not(isnull(rs("mtgysjr"))) and isnull(rs("mtgyshr")) then sel_opt("ģͷ�������")

			if isnull(rs("mtbomr")) and not isnull(rs("mtshjs")) then sel_opt("��ʼģͷBOM")
			if not isnull(rs("mtbomr")) and isnull(rs("mtbomjs")) then sel_opt("����ģͷBOM")

		case "�������"
		If rs("zz")=session("userName") or rs("jgzz")=session("userName") Then'���ṹ�鳤�ܷ���ṹ����
			if isnull(rs("dxjgr")) then sel_opt("��ʼ���ͽṹ")
			if not isnull(rs("dxjgr")) and isnull(rs("dxjgjs")) then sel_opt("�������ͽṹ")
		End If

		If rs("zz")=session("userName") or rs("sjzz")=session("userName") Then'������鳤�ܷ����������
			if isnull(rs("dxsjr")) and not isnull(rs("dxjgjs")) then sel_opt("��ʼ�������")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("�����������")
		End If

			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("��ʼ�������")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("�����������")

'       	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'       	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")

			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼ����BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")

		case "���͸���"
			if isnull(rs("dxsjr")) then sel_opt("��ʼ���͸���")
			if not isnull(rs("dxsjr")) and isnull(rs("dxsjjs")) then sel_opt("�������͸���")

			if isnull(rs("dxshr")) and not isnull(rs("dxsjjs")) then sel_opt("��ʼ�������")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("�����������")

'   	    	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'     	  	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")

			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼ����BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")

		case "���͸���"
			if isnull(rs("dxshr")) then sel_opt("��ʼ���͸���")
			if not isnull(rs("dxshr")) and isnull(rs("dxshjs")) then sel_opt("�������͸���")

'   	    	if not(isnull(rs("dxshjs"))) and isnull(rs("dxgysjr")) then sel_opt("���͹������")
'     	  	if not(isnull(rs("dxgysjr"))) and isnull(rs("dxgyshr")) then sel_opt("���͹������")

			if isnull(rs("dxbomr")) and not isnull(rs("dxshjs")) then sel_opt("��ʼ����BOM")
			if not isnull(rs("dxbomr")) and isnull(rs("dxbomjs")) then sel_opt("��������BOM")
		case else
			response.write(rs("mjxx") & rs("rwlr"))
	end select
%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type="submit" value=" ���� " /></td>
    </tr>
    <input type="hidden" name="lsh" value="<%=rs("lsh")%>" />
  </form>
</table>
<script language="JavaScript" type="text/javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='����')
				objdoc.form_assign.zrr.disabled=true;
			else
				objdoc.form_assign.zrr.disabled=false;
		}
		checkjs();

		function checksubinf(frm)
		{
			if (frm.fplr.value==""){alert("��ѡ���������!"); frm.fplr.focus(); return false;}
			if ((frm.zrr.value=="") && (!frm.zrr.disabled)){alert("��ѡ��������!"); frm.zrr.focus(); return false;}
			return true;
		}
	</script>
<%
end function

function sel_opt(str)
%>
<option value="<%=str%>"><%=str%></option>
<%
end function

function jsdb_assign(rs)
%>
<table border="0" cellpadding="3" cellspacing="0">
  <form action="mtask_assignindb.asp" method="post" name="form_assign" id="form_assign" onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style="font-size:14px;font-weight:bold;">�����ͬ�� <%=rs("hth")%> ��������������:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
		if isnull(rs("sjr")) then sel_opt("��ʼ�����������")
		if not isnull(rs("sjr")) and isnull(rs("sjjssj")) then sel_opt("���������������")

		if isnull(rs("shr")) and not isnull(rs("sjjssj")) then sel_opt("��ʼ�����������")
		if not isnull(rs("shr")) and isnull(rs("shjssj")) then sel_opt("���������������")
%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type="submit" value=" ���� " /></td>
    </tr>
    <input type="hidden" name="hth" value="<%=rs("hth")%>" />
  </form>
</table>
<script language="JavaScript" type="text/javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='����')
				objdoc.form_assign.zrr.disabled=true;
			else
				objdoc.form_assign.zrr.disabled=false;
		}
		checkjs();

		function checksubinf(frm)
		{
			if (frm.fplr.value==""){alert("��ѡ���������!"); frm.fplr.focus(); return false;}
			if ((frm.zrr.value=="") && (!frm.zrr.disabled)){alert("��ѡ��������!"); frm.zrr.focus(); return false;}
			return true;
		}
	</script>
<%
End function

function task_assign_nofinished()			'���з���Ȩ�޵�δ��ɵ�������2013/1/19 10:59
	Dim sqlorder
	sqlorder = " order by jhjssj"
	If LCase(strOrder) = "ddh" Then sqlorder = " order by ddh desc, lsh desc"
	If LCase(strOrder) = "lsh" Then sqlorder = " order by lsh desc"

	If chkable(3) Then
		strSql="select * from [mtask] where (not isnull(fsr)) and isnull(sjjssj) " & sqlorder
	else
		strSql="select * from [mtask] where (zz='"&session("userName")&"' or jgzz='"&session("userName")&"' or sjzz='"&session("userName")&"'or [group]="&session("userGroup")&") and isnull(fsjs) and isnull(sjjssj)" & sqlorder
	End If
	call xjweb.exec(strSql,0)
	rs.open strsql,conn,1,3
	if (rs.eof or rs.bof) then Call TbTopic("��ʱû�д��������������!") : Exit Function
	dim icounter
	Call CutLine()		'��ʾͼ��
	icounter = 1
	sjcount=Rs.recordcount
%>
<%Call TbTopic("������ " & sjcount & " ����������������������")%>
<table width="100%" cellspacing="0" class="xtable">
  <tr>
    <td class="th" width="*">ID</td>
    <td class="th" width="*">��ˮ��</td>
    <td class="th" width="*">��λ��</td>
    <td class="th" width="7%">������</td>
    <td class="th" width="7%">&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;</td>
    <td class="th" width="11%">�ƻ��ṹ���</td>
    <td class="th" width="11%">�ƻ��������</td>
    <td class="th" width="*">����</td>
    <td class="th" width="7%">�ṹ</td>
    <td class="th" width="7%">���</td>
    <td class="th" width="7%">���</td>
    <td class="th" width="*">����</td>
    <td class="th" width="7%">BOM</td>
  </tr>
  <%do while not rs.eof
	'�ṹ\��ƶ�����ʱ��ͼ��15:54 2007-4-1-������
	'���һ�첻�����ڣ�ʵ��Ҫ���������Ӧ��Ϊdatediff("d", �ƻ���ʼ, �ƻ�����)+1
	'���ݿ��˰취�ṹ����=(�������-1)/2����˽ṹ���ڣ��������/2
  Dim jgjgsj, sjjgsj, Tmpsj, ijgsj, isj
If rs("jhkssj")<>"" Then
	Tmpsj=rs("jhkssj")
else
	Tmpsj=rs("rwxdsj")
End If
  jgjgsj=datediff("d", Tmpsj, rs("jhjssj"))/2
  sjjgsj=datediff("d", Tmpsj, rs("jhjssj"))+1-jgjgsj
If IsNull(rs("jhjgsj")) Then
	isj=INT(datediff("d", rs("jhkssj"), rs("jhjssj"))/2)
	ijgsj=dateadd("d",isj,rs("jhkssj"))
else
	isj=rs("jhjssj")
	ijgsj=rs("jhjgsj")
End if
  %>
  <tr>
    <td class="ctd"><%=icounter%></td>
    <td class="ctd"><a href="mtask_assign.asp?s_LSH=<%=rs("LSH")%>"><%=rs("LSH")%></a></td>
    <td class="ctd"><%=rs("DWMC")%></td>
    <td class="ctd"><%=rs("DMMC")%></td>
    <td class="ctd">
    <%If Rs("zz")<>"" Then
    		Response.Write(rs("zz"))
   	ElseIf rs("jgzz")=rs("sjzz") Then
    		Response.Write(rs("jgzz"))
    else
    	Response.Write(rs("jgzz")&"<br>"&rs("sjzz"))
     End If%>
    </td>
    <%if isNull(rs("JHKSSJ")) then%>
    <td class="ctd">&nbsp;/&nbsp;</td>
    <%else%>
    <td class="ctd" alt='�ƻ���ʼ����:<%=rs("jhkssj")%>'><%=rs("jhjgsj")%>&nbsp;</td>
    <%end if%>
    <%if isNull(rs("SJJSSJ")) then%>
    <td class="ctd" alt='��δ���'><%= rs("JHJSSJ") %></td>
    <%else%>
    <td class="ctd" alt='ʵ�ʽ�������:<%=rs("SJJSSJ")%>'><%= rs("JHJSSJ") %></td>
    <%end if%>
    <td class="ctd"><%= rs("MJXX") & rs("RWLR") %></td>
    <%select case rs("mjxx") & rs("rwlr")%>
    <%case "ȫ�����"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%=jgjgsj%>
      <%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <% If not(isnull(rs("gjjgks"))) Then call DisTdjg(rs("gjjgks"),rs("gjjgjs"),ijgsj,rs)%>
    </td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%=sjjgsj%>
      <%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <% If not(isnull(rs("gjsjks"))) Then call DisTdjg(rs("gjsjks"),rs("gjsjjs"),isj,rs)%>
    </td>
    <td class="ctd"><%If not(isnull(rs("mtshr"))) or not(isnull(rs("dxshr"))) Then
      		call distd(rs("mtshks"),rs("mtshjs"),0,rs)
      		call distd(rs("dxshks"),rs("dxshjs"),0,rs)
     		If not(isnull(rs("gjshr"))) Then call distd(rs("gjshks"),rs("gjshjs"),0,rs) End If
     	else
     		call DisTdjg(rs("mtjgshks"),rs("mtjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("dxjgshks"),rs("dxjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("mtsjshks"),rs("mtsjshjs"),isj,rs)
       		call DisTdjg(rs("dxsjshks"),rs("dxsjshjs"),isj,rs)
       		If not(isnull(rs("gjjgr"))) Then call DisTdjg(rs("gjjgshks"),rs("gjjgshjs"),ijgsj,rs) End If
       		If not(isnull(rs("gjsjr"))) Then call DisTdjg(rs("gjsjshks"),rs("gjsjshjs"),isj,rs) End If
      End If%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "ȫ�׸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
      <%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "ȫ�׸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
      <%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
      <%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "ģͷ���"%>
    <td class="ctd"><%call DisTdjg(rs("mtjgks"),rs("mtjgjs"),ijgsj,rs)%>
      <%=jgjgsj%> </td>
    <td class="ctd"><%call DisTdjg(rs("mtsjks"),rs("mtsjjs"),isj,rs)%>
      <%=sjjgsj%> </td>
    <td class="ctd"><%If not(isnull(rs("mtshr"))) Then
      		call distd(rs("mtshks"),rs("mtshjs"),0,rs)
     	else
     		call DisTdjg(rs("mtjgshks"),rs("mtjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("mtsjshks"),rs("mtsjshjs"),isj,rs)
      End If%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
    </td>
    <%case "ģͷ����"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtsjks"),rs("mtsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
    </td>
    <%case "ģͷ����"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("mtshks"),rs("mtshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtgysjks"),rs("mtgysjjs"),0,rs)%>
      <%call distd(rs("mtgyshks"),rs("mtgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("mtbomks"),rs("mtbomjs"),0,rs)%>
    </td>
    <%case "�������"%>
    <td class="ctd"><%call DisTdjg(rs("dxjgks"),rs("dxjgjs"),ijgsj,rs)%>
      <%=jgjgsj%> </td>
    <td class="ctd"><%call DisTdjg(rs("dxsjks"),rs("dxsjjs"),isj,rs)%>
      <%=sjjgsj%> </td>
    <td class="ctd"><%If not(isnull(rs("dxshr"))) Then
      		call distd(rs("dxshks"),rs("dxshjs"),0,rs)
     	else
       		call DisTdjg(rs("dxjgshks"),rs("dxjgshjs"),ijgsj,rs)
       		call DisTdjg(rs("dxsjshks"),rs("dxsjshjs"),isj,rs)
      End If%>
    </td>
    <td class="ctd">
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "���͸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxsjks"),rs("dxsjjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd">
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%case "���͸���"%>
    <td class="ctd">&nbsp;</td>
    <td class="ctd">&nbsp;</td>
    <td class="ctd"><%call distd(rs("dxshks"),rs("dxshjs"),0,rs)%>
    </td>
    <td class="ctd">
      <%call distd(rs("dxgysjks"),rs("dxgysjjs"),0,rs)%>
      <%call distd(rs("dxgyshks"),rs("dxgyshjs"),0,rs)%>
    </td>
    <td class="ctd"><%call distd(rs("dxbomks"),rs("dxbomjs"),0,rs)%>
    </td>
    <%end select%>
  </tr>
  <%
		rs.movenext
		icounter = icounter + 1
	loop
	rs.close
%>
</table>
<%End function

function jsdb_nofinished()			'���з���Ȩ�޵�δ��ɵļ�������������
	Dim sqlorder, icounter
	If LCase(strOrder) = "ddh" or LCase(strOrder) = "lsh" Then
		sqlorder = " order by hth desc"
	else
		sqlorder = " order by jhjssj"
	End If

	strSql="select * from [jsdb] where zz='"&session("userName")&"' and isnull(shjssj)" & sqlorder
'	Set Rs=xjweb.Exec(strSql,1)
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.open strSql,Conn,1,3
	if (Rs.eof or Rs.bof) then Call TbTopic("��ʱû�д����似������������!") : Exit Function
	icounter = 1
	dbcount=Rs.recordcount
	Call TbTopic("������ " & dbcount & " ��������ļ�����������������")
%>
<table width="100%" cellspacing="0" class="xtable">
  <tr>
    <td class="th" width="*">ID</td>
    <td class="th" width="*">������</td>
    <td class="th" width="15%">�ͻ�����</td>
    <td class="th" width="*">��������</td>
    <td class="th" width="7%">&nbsp;&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;��&nbsp;&nbsp;</td>
    <td class="th" width="11%">�ƻ����ʱ��</td>
    <td class="th" width="7%">���</td>
    <td class="th" width="7%">���</td>
  </tr>
  <%do while not rs.eof%>
  <tr>
      <td class="ctd"><%=icounter%></td>
    <td class="ctd"><a href="mtask_assign.asp?s_hth=<%=rs("hth")%>"><%=rs("hth")%></a></td>
    <td class="ctd"><%=rs("khmc")%></td>
    <td class="ctd"><%=rs("rwnr")%></td>
    <td class="ctd"><%=rs("zz")%></td>
    <td class="ctd"><%=rs("jhjssj")%></td>
    <td class="ctd"><%call distd(rs("sjkssj"),rs("sjjssj"),0,rs)%></td>
    <td class="ctd"><%call distd(rs("shkssj"),rs("shjssj"),0,rs)%></td>
</tr>
  <%
		rs.movenext
		icounter = icounter + 1
	loop
	rs.close
%>
</table>
<%end function%>