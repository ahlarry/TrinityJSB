<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtask_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'10:52 2007-1-25-������
Call ChkPageAble("3,4")
CurPage="�������� �� ������Ϣ�����������"
strPage="atask"
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
<table class="xtable" cellspacing="0" cellpadding="2" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd" style="border-right-style:none"><table border="0" cellpadding="2" cellspacing="0" width="100%">
        <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="post" name="frm_searchlsh" id="frm_searchlsh" onsubmit='return searchlsh_true();'>
          <tr>
            <td>&nbsp;&nbsp;���붩����:
              <input tabindex="1" type="text" name="s_ddh" size="15" value="<%=Trim(Request("s_ddh"))%>" />
            </td>
            <td>&nbsp;&nbsp;������ˮ��:
              <input tabindex="1" type="text" name="s_ls" size="15" value="<%=Trim(Request("s_ls"))%>" />
              <input type="submit" value=" �� �� " />
            </td>
          </tr>
        </form>
      </table></td>
    <td class="rtd" style="border-left-style:none"><%Call Taxis()%></td>
  </tr>
  <tr>
    <td class="ctd" height="300" colspan="2"><%Call InfoFix_add()%>
      <%Response.Write(XjLine(1,"100%",web_info(12)))%>
      <%Call InfoFix_nofinished() %>
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
  <option value="ddh" selected="selected">������</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>��ˮ��</option>
  <option value="khmc" <%If strOrder="khmc" Then%>selected<%End If%>>�ͻ�����</option>
</select>
<%
End Function

Function InfoFix_add()
	Dim s_ddh, s_ls, s_time
	s_ddh="" : s_ls="" : s_time=""
	If Trim(Request("s_ddh"))<>"" Then s_ddh=Trim(Request("s_ddh"))
	If Trim(Request("s_ls"))<>"" Then s_ls=Trim(Request("s_ls"))
	If ( s_ddh="" and s_ls="") Then Call TbTopic("������Ҫ���ĵ�������Ķ����Ż���ˮ��!") : Exit Function

'	strSql="select a.*, b.tscs, b.tscs as tscs from [ts_mould] b, [mtask] a where [ddh]='"&s_ddh&"' and ((mjxx='ȫ��'and not(isnull(mttsxxzlr)) and not(isnull(dxtsxxzlr))) or (mjxx='ģͷ'and not(isnull(mttsxxzlr))) or (mjxx='����'and not(isnull(dxtsxxzlr)))) and isnull(xtxxsjjs) and b.tscs>1 and a.lsh=b.lsh"
	strSql="select * from [mtask] a where ([ddh]='"&s_ddh&"' or [lsh]='"&s_ls&"') and isnull(xtxxsjjs) and (not(isNull(mttsdjs)) or not(isNull(dxtsdjs)))"
	set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		If s_ddh<>"" Then
			Call JsAlert("������Ϊ �� " & s_ddh & " �� �������鲻���ڻ��ѽ���!", "InfoFix_add.asp")
		Else
			If s_ls<>"" Then
				Call JsAlert("��ˮ��Ϊ �� " & s_ls & " �� �������鲻���ڻ��ѽ���!", "InfoFix_add.asp")
			End If
		End If
	Else
			Call InfoFixadd(Rs)
	End If
	Rs.Close
End Function

Function InfoFixadd(Rs)
Call TbTopic("���ģ��������Ϣ����������")
%>
<table id="table1" class="ktable" cellspacing="0" cellpadding="3" width="98%" align="center">
  <form id="InfoFix_add" name="InfoFix_add" action="InfoFix_indb.asp?action=add" method="post">
    <tr>
      <th class="rtd" width="15%">��Ŀ����
        </td>
      </th>
      <th class="ltd">��Ŀ����
        </td>
      </th>
    </tr>
    <tr>
      <td class="rtd">������:</td>
      <td class="ltd"><input name="lsh" type="text" disabled="disabled" id="ddh" onchange="FindLsh();" value="<%=Rs("ddh")%>" size="15" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        �����ܷ�:<span id="span_rwzf" style="font-weight:bold;">0</span>�� </td>
    </tr>
    <%
Dim m,sjxs
m=1 : sjxs=0
while not rs.eof
	If Rs("xtxxzlr")<>"" Then
%>
    <tr>
      <td class="rtd">ִ����:</td>
      <td class="ltd"><%=Rs("xtxxzlr")%> &nbsp;&nbsp;��ˮ��:<%=Rs("lsh")%>&nbsp;&nbsp;��ֵϵ��:<%=Rs("xtxxzlxs")%> &nbsp;&nbsp;�����:<%=Round(Rs("mjzf")*Rs("xtxxzlxs"),1)%>&nbsp;&nbsp;�ƻ�����ʱ��:<%=Rs("xtxxjhjs")%>&nbsp;&nbsp;
        <select name="xtzlwc<%=m%>">
          <option></option>
          <option value='�������'>������Ϣ�������</option>
          <input type="hidden" name="zxr<%=m%>" value="<%=Rs("xtxxzlr")%>" />
          <input type="hidden" name="lsh<%=m%>" value="<%=Rs("lsh")%>" />
          <input type="hidden" name="fzxs<%=m%>" value="" />
          <input type="hidden" name="zlxf<%=m%>" value=<%=Round(Rs("mjzf")*Rs("xtxxzlxs"),1)%> />
        </select>
      </td>
    </tr>
    <%else%>
    <tr>
      <td class="rtd">ִ����:</td>
      <td class="ltd"><select name="zxr<%=m%>" onchange="fzhi(<%=m%>,1);">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;��ˮ��:<%=Rs("lsh")%> &nbsp;&nbsp;ģ�߷�:<%=Rs("mjzf")%>&nbsp;&nbsp;&nbsp;&nbsp;��ֵϵ��:
        <input name="fzxs<%=m%>" type="text" onchange="fzhi(<%=m%>,2);" value="0.08" size="5" />
        &nbsp;&nbsp;�����:<span id="span_zldf<%=m%>" style="font-weight:bold;">0</span>��
        <input type="hidden" name="mjf<%=m%>" value="<%=Rs("mjzf")%>" />
        <input type="hidden" name="lsh<%=m%>" value="<%=Rs("lsh")%>" />
        <input type="hidden" name="zlxf<%=m%>" value="0" />
      </td>
    </tr>
    <%
   	sjxs=1
	End If
  rs.movenext
  m=m+1
  wend%>
    <input type="hidden" name="Num" value="<%=m-1%>" />
    <%If sjxs=1 Then%>
    <tr>
      <td class="rtd">�ƻ�����ʱ��</td>
      <td class="ltd"><select id="psy" name="psy" onchange='addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);'>
          <%for i = year(now)-1 to year(now) + 3%>
          <%if i = year(now) then%>
          <option value='<%=i%>' selected="selected"><%=i%></option>
          <%else%>
          <option value='<%=i%>'><%=i%></option>
          "
          <%end if%>
          <%next%>
        </select>
        ��
        <select id="psm" name="psm" onchange='addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);'>
          <%for i = 1 to 12%>
          <%if i = month(now) then%>
          <option value='<%=i%>' selected="selected"><%=i%></option>
          <%else%>
          <option value='<%=i%>'><%=i%></option>
          <%end if%>
          <%next%>
        </select>
        ��
        <select id="psd" name="psd">
          <%for i = 1 to 31%>
          <%if isdate(year(now) & "-" & month(now) & "-" & i) then%>
          <%if i = day(now) then%>
          <option value='<%=i%>' selected="selected"><%=i%></option>
          <%else%>
          <option value='<%=i%>'><%=i%></option>
          <%end if%>
          <%end if%>
          <%next%>
        </select>
        �� </td>
    </tr>
    <%End If%>
    <tr>
      <td colspan="2" align="center"><input name="submit" type="submit" value=" �� ȷ �� �� " />
      </td>
    </tr>
  </form>
</table>
<%
End Function	'End InfoFixadd()

Function InfoFix_nofinished()			'���з���Ȩ�޵�δ��ɵĵ�������
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter,TotalCount,s_ddh, s_ls, sqlorder
	absPageNum = 0
	RecordPerPage = 20
	iCounter = 1
	s_ddh="" : s_ls=""
	sqlorder = " order by ddh desc, lsh desc"
	If LCase(strOrder) = "khmc" Then sqlorder = " order by dwmc"
	If LCase(strOrder) = "lsh" Then sqlorder = " order by lsh desc"

	If Trim(Request("s_ddh"))<>"" Then s_ddh=Trim(Request("s_ddh"))
	If Trim(Request("s_ls"))<>"" Then s_ls=Trim(Request("s_ls"))

'	strSql="select a.*, b.tscs, b.tscs as tscs from [ts_mould] b, [mtask] a where not(mjjs) and ((mjxx='ȫ��'and not(isnull(mttsxxzlr)) and not(isnull(dxtsxxzlr))) or (mjxx='ģͷ'and not(isnull(mttsxxzlr))) or (mjxx='����'and not(isnull(dxtsxxzlr)))) and isnull(xtxxsjjs) and b.tscs>1 and a.lsh=b.lsh order by a.ddh desc"
	If (s_ddh<>"" or s_ls<>"")Then
		strSql="select * from [mtask] where ([ddh]='"&s_ddh&"' or [lsh]='"&s_ls&"') and not(mjjs) and isnull(xtxxsjjs) and (not(isNull(mttsdjs)) or not(isNull(dxtsdjs)))" & sqlorder
	else
		strSql="select * from [mtask] where not(mjjs) and isnull(xtxxsjjs) and (not(isNull(mttsdjs)) or not(isNull(dxtsdjs)))" & sqlorder
	End If

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
<%Call TbTopic("���� "& rs.recordcount &" �״�����ĸ�������")%>
<table width="95%" cellspacing="0" cellpadding="2" class="xtable" align="center">
  <tr>
    <th class="th">id</th>
    <th class="th">������</th>
    <th class="th">��ˮ��</th>
    <th class="th">��λ����</th>
    <th class="th">��������</th>
    <th class="th">��������</th>
    <th class="th" width="*">������Ϣ����</th>
    <th class="th" width="*">ģ��������Ϣ����</th>
    <th class="th" width="*">��������ƻ�����</th>
  </tr>
  <% for absrecordnum = 1 to recordperpage %>
  <tr>
    <td class="ctd"><%=icounter %></td>
    <td class="ctd"><a href="InfoFix_add.asp?s_ddh=<%=rs("ddh")%>"><%= rs("ddh")%></td>
    <td class="ctd"><a href="InfoFix_add.asp?s_ls=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("dwmc")%></td>
    <td class="ctd"><%=rs("dmmc")%></td>
    <td class="ctd"><%=rs("jsdb")%></td>
    <%select case rs("mjxx")%>
    <%case "ȫ��"%>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
      <%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%>
    </td>
    <td class="ctd"><%=Rs("xtxxjhjs")%>&nbsp;</td>
    <%case "ģͷ"%>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%></td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%></td>
    <td class="ctd"><%=Rs("xtxxjhjs")%>&nbsp;</td>
    <%case "����"%>
    <td class="ctd"><%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%></td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%>
    </td>
    <td class="ctd"><%=Rs("xtxxjhjs")%>&nbsp;</td>
    <%end select%>
  </tr>
  <%rs.movenext%>
  <%if rs.eof then%>
  <%exit for%>
  <%end if%>
  <%icounter = icounter - 1%>
  <%next%>
</table>
<table width="95%" cellpadding="2" cellspacing="0" border="0" align="center">
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
%>
<script language="JavaScript" type="text/javascript">
 function fzhi(x,y)		//��������ϵ���͹�ͬ���������ȷ������ϵ������ֵ
{
	var tmpfz=0;
	eval("document.all.span_zldf" + x + ".innerHTML=Math.round(document.all.mjf" + x + ".value*document.all.fzxs" + x + ".value*100)/100.0;");	//��ʾ����ˮ�ŷ�ֵ
	eval("document.all.zlxf" + x + ".value=Math.round(document.all.mjf" + x + ".value*document.all.fzxs" + x + ".value*100)/100.0;");
	for (i=1; i<=document.all.Num.value; i++)
  {
    if (tmpfz == 0)
    {
    	tmpfz=eval("document.all.zlxf" + i + ".value;");
    }
    else
    {
    	tmpfz=Math.round((Number(tmpfz) + Number(eval("document.all.zlxf" + i + ".value;")))*100)/100.0;
    }
  }
    document.all.span_rwzf.innerHTML=tmpfz;
	if (eval("document.all.zxr" + Number(x+1) + ".value==''") && (y==1))
      {
	eval("document.all.zxr" + Number(x+1) + ".selectedIndex=document.all.zxr"+x+".selectedIndex;");
      }
	fzhi(Number(x+1),y);
}
</script>
