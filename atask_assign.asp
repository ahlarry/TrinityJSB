<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
Call ChkPageAble("3,4,6")
CurPage="�������� �� �����������"					'ҳ�������λ��( ��������� �� ���������)
strPage="atask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <tr>
    <td class="ctd" style="border-right-style:none"><%Call SearchLsh()%></td>
    <td class="rtd" style="border-left-style:none"><%Call Taxis()%></td>
  </tr>
  <Tr>
    <Td class=ctd height=300 colspan="2"><%Call ataskAssign()%>
      <%Response.Write(XjLine(10,"100%",""))%>
      <%Response.Write(XjLine(1,"100%",web_info(12)))%>
      <%Call atask_nofinished() %>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function Taxis()
Dim strO
strO="1"
%>
����:
<select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;order=&quot; + this.value);'>
  <option value="jhjssj" selected="selected">���Ե��ƻ����ʱ��</option>
  <option value="ddh" <%If strOrder="ddh" Then%>selected<%End If%>>������</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>��ˮ��</option>
</select>
<%
End Function

'�������������,һ��Ϊ���鳤����(�����Ե�),��һ��Ϊ�����鳤����(���Լ�������Ϣ����)
Function ataskAssign()
	Dim s_lsh
	s_lsh=""
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("��������丨���������ˮ��!") : Exit Function

	'strSql="select a.*, b.*, a.lsh as lsh from [mtask] a, [mtask] b where a.lsh='"&s_lsh&"' and a.lsh=b.lsh and (not isnull(sjjssj)) and not mjjs"
	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("��ˮ�� ��" & s_lsh & "�� �����鲻����!","atask_assign.asp") : Exit Function
	ElseIf IsNull(Rs("fsjs")) Then
		Call JsAlert("��ˮ�� ��" & s_lsh & "�� ���������������!","atask_assign.asp") : Exit Function
	ElseIf Rs("mjjs") Then
		Call JsAlert("��ˮ�� ��" & s_lsh & "�� �������Ѿ�ȫ�����!","atask_assign.asp") : Exit Function
	Else
		Select Case Rs("mjxx")
			Case "ȫ��"
				If not(isnull(rs("mttsdjs"))) and not(isnull(rs("dxtsdjs")))	then
					call group5_assign(rs)
				else
					'response.write rs("group")
					call group_assign(rs)
				end if
			Case "ģͷ"
				if not(isnull(rs("mttsdjs"))) then
					call group5_assign(rs)
				else
					call group_assign(rs)
				end if
			Case "����"
				if not(isnull(rs("dxtsdjs"))) then
					call group5_assign(rs)
				else
					call group_assign(rs)
				end if
		End Select
	End If
	Rs.Close
End Function

Function mtask_info(rs)
	Call mtask_fewinfo(rs)
	Response.Write(xjLine(4, "100%", ""))%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%" align="center">
  <tr>
    <td class="rtd" width="15%">���ڵ���</td>
    <td colspan="3" class="ltd" width="35%"><%if rs("cnts") then%>
      ��
      <%else%>
      &nbsp;/
      <%end if%></td>
    <td class="rtd" width="15%">�������</td>
    <% If Rs("cnts") Then%>
    <%If Not(isnull(Rs("tslb"))) Then%>
    <td class="ltd"><%=Rs("tslb")%></td>
    <%Else%>
    <td class="ltd">&nbsp;/</td>
    <%End If%>
    <%Else%>
    <%If Rs("beit") Then%>
    <td class="ltd">����</td>
    <%Else%>
    <td class="ltd">&nbsp;/</td>
    <%End If%>
    <%End If%>
  </tr>
</table>
	<%Response.Write(xjLine(4, "100%", ""))
	Call atask_userinfo(rs)
	Response.Write(xjLine(8, "100%", ""))
	Response.Write(xjLine(1, "100%", "#00659c"))
	Response.Write(xjLine(8, "100%", ""))
End Function

Function group_assign(rs)
	If rs("group") = Session("userGroup") Or Rs("zz")=Session("UserName") Or Rs("jgzz")=Session("UserName") Or Rs("sjzz")=Session("UserName") Then
		Call mtask_info(Rs)
	%>
<table border=0 cellpadding=3 cellspacing=0 align="center">
  <form name="form_assign" action="atask_assignindb.asp" method=post onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style=font-size:14px;font-weight:bold;>������ˮ�� <%=rs("lsh")%> ��������:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
				select case rs("mjxx")
					case "ȫ��"
						if isnull(rs("mttsdr")) then sel_opt("��ʼģͷ���Ե�")
						if isnull(rs("dxtsdr")) then sel_opt("��ʼ���͵��Ե�")
						if isnull(rs("mttsdr")) and isnull(rs("dxtsdr")) then sel_opt("��ʼȫ�׵��Ե�")

						if (not isnull(rs("mttsdr"))) and isnull(rs("mttsdjs")) then sel_opt("����ģͷ���Ե�")
						if not isnull(rs("dxtsdr")) and isnull(rs("dxtsdjs")) then sel_opt("�������͵��Ե�")
						if (rs("mttsdr")=rs("dxtsdr")) and isnull(rs("mttsdjs")) and isnull(rs("dxtsdjs")) then sel_opt("����ȫ�׵��Ե�")

					case "ģͷ"
						if isnull(rs("mttsdr")) then sel_opt("��ʼģͷ���Ե�")
						if (not isnull(rs("mttsdr"))) and isnull(rs("mttsdjs")) then sel_opt("����ģͷ���Ե�")

					case "����"
						if isnull(rs("dxtsdr")) then sel_opt("��ʼ���͵��Ե�")
						if not isnull(rs("dxtsdr")) and isnull(rs("dxtsdjs")) then sel_opt("�������͵��Ե�")
					case else
						response.write(rs("mjxx"))
				end select
				%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
          <option value="TT����Ա">TT����Ա</option>
          <option value="TB����Ա">TB����Ա</option>
        </select>
      </td>
      <td><input type=submit value=" ���� "></td>
    </tr>
    <input type=hidden name=lsh value=<%=rs("lsh")%>>
  </form>
</table>
<%
		call atask_js()
	Else
		Call JsAlert("��������鳤�� " & rs("zz") & rs("jgzz") &"(�ṹ)��"& rs("sjzz") &"(���) ! ����ϵ�鳤�����������!","")
	End If
end function

Function group5_assign(rs)
	If Session("userGroup") <> 5 Then
		Call JsAlert("����ϵ�����鳤���д˸����������!","atask_assign.asp")
	Else
		Call mtask_info(Rs)
	%>
<table border=0 cellpadding=3 cellspacing=0 align="center">
  <form name="form_assign" action="atask_assignindb.asp" method=post onsubmit="return(checksubinf(this));">
    <tr>
      <td><font style=font-size:14px;font-weight:bold;>������ˮ�� <%=rs("lsh")%> ��������:</font></td>
      <td><select name="fplr" onchange="checkjs();">
          <%
					select case rs("mjxx")
						case "ȫ��"
							if isnull(rs("mttsr")) and not isnull(rs("mttsdjs")) and isnull(rs("dxtsr")) and not isnull(rs("dxtsdjs")) then sel_opt("��ʼȫ�׵���")
							if isnull(rs("mttsr")) and not isnull(rs("mttsdjs")) then sel_opt("��ʼģͷ����")
							if isnull(rs("dxtsr")) and not isnull(rs("dxtsdjs")) then sel_opt("��ʼ���͵���")

							if (rs("mttsr")=rs("dxtsr")) and isnull(rs("mttsjs"))  and isnull(rs("dxtsjs")) then
								sel_opt("����ȫ�׵���")
								sel_opt("ȫ�׳��ڳ���")
								sel_opt("ȫ�׳��⾫��")
								sel_opt("ȫ��Ԥ���ջ����")
								sel_opt("ȫ����������")
							End If
							if not isnull(rs("mttsr")) and isnull(rs("mttsjs")) then
								sel_opt("����ģͷ����")
								sel_opt("ģͷ���ڳ���")
								sel_opt("ģͷ���⾫��")
								sel_opt("ģͷԤ���ջ����")
								sel_opt("ģͷ��������")
							End If
							if not isnull(rs("dxtsr")) and isnull(rs("dxtsjs")) then
								sel_opt("�������͵���")
								sel_opt("���ͳ��ڳ���")
								sel_opt("���ͳ��⾫��")
								sel_opt("����Ԥ���ջ����")
								sel_opt("������������")
							End If
							if isnull(rs("mttsxxzlr")) and not isnull(rs("mttsjs")) then sel_opt("��ʼģͷ������Ϣ����")
							if isnull(rs("dxtsxxzlr")) and not isnull(rs("dxtsjs")) then sel_opt("��ʼ���͵�����Ϣ����")
							if isnull(rs("mttsxxzlr")) and not isnull(rs("mttsjs")) and isnull(rs("dxtsxxzlr")) and not isnull(rs("dxtsjs")) then sel_opt("��ʼȫ�׵�����Ϣ����")

							if not isnull(rs("mttsxxzlr")) and isnull(rs("mttsxxzljs")) then sel_opt("����ģͷ������Ϣ����")
							if not isnull(rs("dxtsxxzlr")) and isnull(rs("dxtsxxzljs")) then sel_opt("�������͵�����Ϣ����")
							if (rs("mttsxxzlr")=rs("dxtsxxzlr")) and isnull(rs("mttsxxzljs")) and isnull(rs("dxtsxxzljs")) then sel_opt("����ȫ�׵�����Ϣ����")
						case "ģͷ"
							if isnull(rs("mttsr")) and not isnull(rs("mttsdjs")) then sel_opt("��ʼģͷ����")
							if not isnull(rs("mttsr")) and isnull(rs("mttsjs")) then
								sel_opt("����ģͷ����")
								sel_opt("ģͷ���ڳ���")
								sel_opt("ģͷ���⾫��")
								sel_opt("ģͷԤ���ջ����")
								sel_opt("ģͷ��������")
							End If
							if isnull(rs("mttsxxzlr")) and not isnull(rs("mttsjs")) then sel_opt("��ʼģͷ������Ϣ����")
							if not isnull(rs("mttsxxzlr")) and isnull(rs("mttsxxzljs")) then sel_opt("����ģͷ������Ϣ����")

						case "����"
							if isnull(rs("dxtsr")) and not isnull(rs("dxtsdjs")) then sel_opt("��ʼ���͵���")
							if not isnull(rs("dxtsr")) and isnull(rs("dxtsjs")) then
								sel_opt("�������͵���")
								sel_opt("���ͳ��ڳ���")
								sel_opt("���ͳ��⾫��")
								sel_opt("����Ԥ���ջ����")
								sel_opt("������������")
							End If
							if isnull(rs("dxtsxxzlr")) and not isnull(rs("dxtsjs")) then sel_opt("��ʼ���͵�����Ϣ����")
							if not isnull(rs("dxtsxxzlr")) and isnull(rs("dxtsxxzljs")) then sel_opt("�������͵�����Ϣ����")
						case else
							response.write(rs("mjxx"))
					end select
				%>
        </select>
        <select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </select>
      </td>
      <td><input type=submit value=" ���� "></td>
    </tr>
    <input type=hidden name=lsh value=<%=rs("lsh")%>>
  </form>
</table>
<%
		call atask_js()
	End If
End Function

Function atask_nofinished()			'���з���Ȩ�޵�δ��ɵĵ�������
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter,TotalCount, sqlorder
	absPageNum = 0
	RecordPerPage = 40
	iCounter = 1
	sqlorder = " order by jhjssj"
	If LCase(strOrder) = "ddh" Then sqlorder = " order by ddh desc, lsh desc"
	If LCase(strOrder) = "lsh" Then sqlorder = " order by lsh desc"

	strSql="select * from [mtask] where not(isnull(fsjs)) and not(mjjs) and ((mjxx='ȫ��' and (isnull(mttsdjs) or isnull(dxtsdjs)) and ([group]="&session("userGroup")&" Or zz='"&Session("userName")&"' Or jgzz='"&Session("userName")&"' Or sjzz='"&Session("userName")&"')) or (mjxx='ģͷ' and isnull(mttsdjs) and ([group]="&session("userGroup")&" Or zz='"&Session("userName")&"' Or jgzz='"&Session("userName")&"' Or sjzz='"&Session("userName")&"')) or (mjxx='����' and isnull(dxtsdjs) and ([group]="&session("userGroup")&" Or zz='"&Session("userName")&"' Or jgzz='"&Session("userName")&"' Or sjzz='"&Session("userName")&"')) or ((mjxx='ȫ��' and (not isnull(mttsdjs)) and (not isnull(dxtsdjs)) and "&session("userGroup")&"=5) or (mjxx='ģͷ' and (not isnull(mttsdjs)) and "&session("userGroup")&"=5) or (mjxx='����' and (not isnull(dxtsdjs)) and "&session("userGroup")&"=5) ))" & sqlorder
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,1,3
	If (Rs.Eof Or Rs.Bof) Then
		Call TbTopic("��ʱû�д����丨������!") : Exit Function
	End if
	Rs.PageSize = RecordPerPage
	TotalCount=Rs.RecordCount

	If Trim(Request("iPage")) <> ""  Then
		If IsNumeric(Trim(Request("iPage"))) Then
			If Trim(Request("iPage")) <= 0 Then
				absPageNum = 1
			ElseIf CLng(Trim(Request("iPage"))) > Rs.PageCount Then
				absPageNum = Rs.PageCount
			Else
				absPageNum = CLng(Trim(Request("iPage")))
			End If
		Else
			If Request("iCurPage") <> "" Then
				absPageNum = CLng(Request("iCurPage"))
			Else
				absPageNum = 1
			End If
		End If
	Else
		If Request("iCurPage") <> "" Then
			absPageNum = CLng(Request("iCurPage"))
		Else
			absPageNum = 1
		End If
	End If

	If absPageNum > Rs.PageCount then absPageNum = Rs.PageCount
	rs.absolutepage = absPageNum
	icounter=totalcount-(abspagenum-1)*recordperpage
%>
<%Call TbTopic("������ "& rs.recordcount &" �״�����ĸ�������")%>
<table width="95%" cellspacing=0 cellpadding=2 class=xtable align="center">
  <tr>
    <th class="th">id</th>
    <th class="th">������</th>
    <th class="th">��ˮ��</th>
    <th class="th">��λ����</th>
    <th class="th">��������</th>
    <th class="th">�鳤</th>
    <th class="th">��������</th>
    <th class=th width=*>���Ե�</th>
    <th class=th width=*>����</th>
    <th class=th width=*>��������</th>
  </tr>
  <% for absrecordnum = 1 to recordperpage %>
  <tr>
    <td class="ctd"><%=icounter %></td>
    <td class="ctd"><%= rs("ddh")%></td>
    <td class="ctd"><a href=atask_assign.asp?s_lsh=<%=rs("lsh")%>><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("dwmc")%></td>
    <td class="ctd"><%=rs("dmmc")%></td>
    <td class="ctd"><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(j)��"&rs("sjzz")&"(s)")%></td>
    <td class="ctd"><%= rs("mjxx") & rs("rwlr") %></td>
    <%select case rs("mjxx")%>
    <%case "ȫ��"%>
    <td class="ctd"><%call distd(rs("mttsdks"),rs("mttsdjs"),-20,rs)%>
      <%=DATEADD("d",20,rs("jhjssj"))%>
      <%call distd(rs("dxtsdks"),rs("dxtsdjs"),-20,rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsks"),rs("mttsjs"),rs)%>
      <%call distd2(rs("dxtsks"),rs("dxtsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
      <%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <%case "ģͷ"%>
    <td class="ctd"><%call distd(rs("mttsdks"),rs("mttsdjs"),-20,rs)%>
      <%=DATEADD("d",20,rs("jhjssj"))%> </td>
    <td class="ctd"><%call distd2(rs("mttsks"),rs("mttsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
    </td>
    <%case "����"%>
    <td class="ctd"><%=DATEADD("d",20,rs("jhjssj"))%>
      <%call distd(rs("dxtsdks"),rs("dxtsdjs"),-20,rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("dxtsks"),rs("dxtsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <%end select%>
  </tr>
  <%rs.movenext%>
  <%if rs.eof then%>
  <%exit for%>
  <%end if%>
  <%icounter = icounter - 1%>
  <%next%>
</table>
<table width="95%" cellpadding=2 cellspacing=0 border=0 align="center">
  <tr>
    <td align=left> ���������� <%=rs.recordcount%> ��&nbsp;&nbsp;
      ÿҳ <%=rs.pagesize%> ��&nbsp;&nbsp;
      �� <%=Rs.PageCount%> ҳ&nbsp;&nbsp;
      ��ǰΪ�� <%=absPageNum%> ҳ </td>
    <td align=right> ��
      <%
				if absPageNum > 1 then
					response.write("<a href="&Request.ServerVariables("script_name")&"?ipage="&(abspagenum-1)&" alt='��һҳ'> ��</a>&nbsp;&nbsp;")
				end if
				Dim iStart,iEnd
				if absPageNum < 4 then
					iStart = 1
				else
					iStart = absPageNum - 3
				end if
				if absPageNum < Rs.PageCount - 3 then
					iEnd = absPageNum + 3
				else
					iEnd = Rs.PageCount
				end if
				for i = iStart to iEnd
					if i = absPageNum then
						response.write("&nbsp;<font style=font-size:11pt;><b>"&  i & "</b></font>&nbsp;")
					else
						response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&i&">" & i & "</a>&nbsp;")
					end if
				next
				if absPageNum < Rs.PageCount then
					response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&(absPageNum+1)&" alt='��һҳ'> �� </a>&nbsp;")
				end if
			%>
      ��
      ��ת��:
      <select name="ipage" onchange='location.href("<%=Request.ServerVariables("script_name")%>?ipage=" + this.value+"");'>
        <%for i=1 to Rs.PageCount%>
        <%if i = absPageNum then%>
        <option value=<%=i%> selected>�� <%=i%> ҳ</option>
        <%else%>
        <option value=<%=i%>>�� <%=i%> ҳ</option>
        <%end if%>
        <%next%>
      </select>
    </td>
  </tr>
</table>
<%
	rs.close
end function

function sel_opt(str)
%>
<option value="<%=str%>"><%=str%></option>
<%
end function

function atask_js()
%>
<script language="javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='��ʼ')
				objdoc.form_assign.zrr.disabled=false;
			else
				objdoc.form_assign.zrr.disabled=true;
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
%>
