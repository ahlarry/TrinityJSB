<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble("3,4")
If Session("userGroup") <> 5 Then Call JsAlert("����ϵ�����鳤���д˸����������!","atask.asp")
CurPage="�������� �� �޸ĵ��������ֵϵ��"					'ҳ�������λ��( ��������� �� ���������)
strPage="atask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, s_lsh
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder
s_lsh=request("s_lsh")

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
    <td class="ctd" height="300" colspan="2"><%If s_lsh="" Then
    		Call atask_nofinished()
    	Else
    		Call ataskAssign()
    	End If%>
      <%Response.Write(XjLine(10,"100%",""))%>
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
  <option value="jssj" selected="selected">���ʱ��</option>
  <option value="zrr" <%If strOrder="zrr" Then%>selected<%End If%>>������</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>��ˮ��</option>
</select>
<%
End Function

Function ataskAssign()
	Dim strXs
	i=1
	If s_lsh="" Then Call TbTopic("�������޸�ϵ������ˮ��!") : Exit Function
	strSql="select * from [mantime] where (rwlr='ȫ�׵��Ժϸ�' or rwlr='ģͷ���Ժϸ�' or rwlr='���͵��Ժϸ�' or rwlr like '%����%' or rwlr like '%����%'  or rwlr like '%����%') and lsh='"&s_lsh&"'"
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.open strSql,Conn,1,3

Call TbTopic("�޸���ˮ�� " & s_lsh & " ���������ֵϵ��")%>
<table width="95%" cellspacing="0" cellpadding="2" class="xtable">
  <tr>
    <th class="th">id</th>
    <th class="th">��ˮ��</th>
    <th class="th">��������</th>
    <th class="th" width="*">������</th>
    <th class="th" width="*">�����ֵ</th>
    <th class="th" width="*">ʵ�ʷ�ֵ</th>
    <th class="th" width="*">ϵ��</th>
    <th class="th" width="*">���ʱ��</th>
    <th class="th" width="*">����޸�ʱ��</th>
  </tr>
  <%do while not Rs.Eof %>
  <tr>
    <td class="ctd"><%=i %></td>
    <td class="ctd"><%=rs("lsh")%></td>
    <td class="ctd"><%=rs("rwlr") %></td>
    <td class="ctd"><%=rs("zrr") %></td>
    <td class="ctd"><%=rs("rwfz") %>&nbsp;</td>
    <td class="ctd"><%=rs("fz") %></td>
    <td class="ctd"><%=rs("jc") %></td>
    <td class="ctd"><%=xjDate(rs("jssj"),1)%></td>
    <td class="ctd" alt="<%=Task_Ts(rs("bz"))%>"><%=xjDate(rs("xgsj"),1)%>&nbsp;</td>
  </tr>
  <%
  strXs=rs("jc")
  rs.movenext
  i = i + 1
  loop
%>
</table>
<table width="95%" cellspacing="0" cellpadding="2">
  <form id=frm_cha name=frm_cha action=atask_xsindb.asp method=post onSubmit='return tscheckinf();'>
    <tr >
      <td height="30" align="center"><input type="hidden" name="yxs" value="<%=strXs%>">
        <input type="hidden" name="ylsh" value="<%=s_lsh%>"></td>
    </tr>
    <tr >
      <td align="center">�������µ�ϵ��:
        <input type="text" id="newxs" name="newxs" value="<%=strXs%>" onkeypress="javascript:validationNumber(this, 'u_float', 10, txtFzMsg);" />
        <input type="submit" value=" �� �� �� �� " />
        <SPAN id="txtFzMsg"></td>
    </tr>
  </form>
</table>
<%
	rs.close
End Function

Function atask_nofinished()			'�����޸�Ȩ�޵�����ɵĵ�������
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter,TotalCount, sqlorder, dtend
	absPageNum = 0
	RecordPerPage = 20
	iCounter = 1
	sqlorder = " order by jssj desc"
	If LCase(strOrder) = "zrr" Then sqlorder = " order by zrr desc, lsh desc"
	If LCase(strOrder) = "lsh" Then sqlorder = " order by lsh desc"

	dtend=(dateadd("m",-1,now()))
	dtend=xjDate(year(dtend)&"��"&month(dtend)&"��1��",1)
	strSql="select * from [mantime] where (rwlr='ȫ�׵��Ժϸ�' or rwlr='ģͷ���Ժϸ�' or rwlr='���͵��Ժϸ�' or rwlr like '%����%' or rwlr like '%����%'  or rwlr like '%����%')  and datediff('m',jssj,'"&dtend&"')<=0" & sqlorder
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,1,3
	If (Rs.Eof Or Rs.Bof) Then
		Call TbTopic(dtend&"�Ժ�û�п����޸ĵĵ��Է�ֵϵ��!") : Exit Function
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
	i=1
%>
<%Call TbTopic("������ "& rs.recordcount &" ��"&dtend&"�Ժ����ĵ��Է�ֵ�����޸�ϵ��")%>
<table width="95%" cellspacing="0" cellpadding="2" class="xtable">
  <tr>
    <th class="th">id</th>
    <th class="th">��ˮ��</th>
    <th class="th">��������</th>
    <th class="th" width="*">������</th>
    <th class="th" width="*">�����ֵ</th>
    <th class="th" width="*">ʵ�ʷ�ֵ</th>
    <th class="th" width="*">ϵ��</th>
    <th class="th" width="*">���ʱ��</th>
    <th class="th" width="*">�޸�ʱ��</th>
  </tr>
  <% for absrecordnum = 1 to recordperpage %>
  <tr>
    <td class="ctd"><%=i %></td>
    <td class="ctd"><a href="atask_changexs.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("rwlr") %></td>
    <td class="ctd"><%=rs("zrr") %></td>
    <td class="ctd"><%=rs("rwfz") %>&nbsp;</td>
    <td class="ctd"><%=rs("fz") %></td>
    <td class="ctd"><%=rs("jc") %></td>
    <td class="ctd"><%=xjDate(rs("jssj"),1)%></td>
    <td class="ctd" alt="<%=Task_Ts(rs("bz"))%>"><%=xjDate(rs("xgsj"),1)%>&nbsp;</td>
  </tr>
  <%rs.movenext
  if rs.eof then
  	exit for
  end if
  i = i + 1
  next%>
</table>
<table width="95%" cellpadding="2" cellspacing="0" border="0">
  <tr>
    <td align="left"> ���������� <%=rs.recordcount%> ��&nbsp;&nbsp;
      ÿҳ <%=rs.pagesize%> ��&nbsp;&nbsp;
      �� <%=Rs.PageCount%> ҳ&nbsp;&nbsp;
      ��ǰΪ�� <%=absPageNum%> ҳ </td>
    <td align="right"> ��
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
      <select name="ipage" onchange='location.href(&quot;<%=Request.ServerVariables("script_name")%>?ipage=&quot; + this.value+&quot;&quot;);'>
        <%for i=1 to Rs.PageCount%>
        <%if i = absPageNum then%>
        <option value="<%=i%>" selected="selected">�� <%=i%> ҳ</option>
        <%else%>
        <option value="<%=i%>">�� <%=i%> ҳ</option>
        <%end if%>
        <%next%>
      </select>
    </td>
  </tr>
</table>
<%
	rs.close
end function

Function Task_Ts(mystr)
    dim strnew
	if instr(mystr,"||")>0 then
		strnew=split(mystr,"||")
    	for i=0 to ubound(strnew)
			Task_Ts=Task_Ts&strnew(i)&"<br>"
		next
	else
		Task_Ts=mystr
	end if
End Function
%>
<script language="javascript">
function tscheckinf()
{
	var objdm=document.frm_cha
	if (objdm.newxs.value==""){alert("ϵ������Ϊ��!"); objdm.newxs.focus(); return false;}
	return true;
}
</script>
