<!--#include file="include/conn.asp"-->
<!--#include file="include/page/ftask_dbinf.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->

<%
Call ChkPageAble("3,10")
CurPage="�������� �� �����������"
strPage="ftask"
Call FileInc(0, "js/ftask.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<table  class="xtable" cellspacing="0" cellpadding="0" width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call Searchxldh()%>
    </td>
  </tr>
  <Tr>
    <Td height=300 class=ctd><%Call NewOrChange()%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</table>
<%
End Sub
Function NewOrChange()
	Dim s_xldh
	s_xldh=""
	If Trim(Request("s_xldh"))<>"" Then s_xldh=Trim(Request("s_xldh"))
	If s_xldh="" Then Call ftaskAdd() : Exit Function
	strSql="Select top 1 * from [ftask] where XLDH='"&s_xldh&"' order by id desc"

	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("������ " & s_xldh & " ������������!","ftask_add.asp")
	Else
		Call ftaskCadd(Rs)
	End If

End Function

Function rwlr_change(i)
         dim mystr,mystr1
		 mystr=rs("rwlr")
	     mystr=split(mystr,"||")
	     If i > ubound(mystr) Then
	     	mystr1=""
	     	rwlr_change=mystr1
	     Else
	   	 	 mystr1=mystr(i)
			 mystr1=split(mystr1,":")
			 rwlr_change=mystr1(1)
		 End If
End Function
function ftaskCadd(rs)
%>
<% Call TbTopic("��Ӹ�����������")%>
	<%call rwlr_change(i)%>
	<form id="frm_ftask" name="ftask_add" action="ftask_indb.asp?action=add" method="post" onsubmit='return checkinf();'>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%" align="center">

	<tr>
		<th class=th height=25>��Ŀ����</th>
		<th class=th>��Ŀ����</th>
	</tr>
	<tr>
		<td class=rtd width="20%">��������</td>
		<td class=ltd>
			<select name="rwlx"><option value=<%=rs("rwlx")%>><%=rs("rwlx")%></option>
			</select>
		</td>
	</tr>
	<tr>
		<td class="rtd">������</td>
        <td class="ltd"><input type="text" name="xldh" size="30"  value=<%=rs("xldh")%>>
         <font color="#FF0000">����</font> </td>
    </tr>
    <tr>
      <td class="rtd">�û���λ</td>
      <td class="ltd"><input type="text" name="yhdw" size="30" value=<%=rwlr_change(0)%>></td>
    </tr>
	<tr>
      <td class="rtd">ģ������</td>
      <td class="ltd"><input type="text" name="mjmc" size="30" value=<%=rwlr_change(1)%>></td>
    </tr>
	<tr>
      <td class="rtd">����С��</td>
      <td class="ltd"><input type="text" name="xlxh" size="30" value=<%=rwlr_change(2)%>></td>
    </tr>
	<tr>
		<td class="rtd">ԭ��ˮ��</td>
        <td class="ltd"><input type="text" name="ylsh" size="30"  value=<%=rwlr_change(6)%>>
         <font color="#FF0000">����</font> </td>
    </tr>
	<tr>
      <td class="rtd">�������������ԭ��</td>
      <td class="ltd"><textarea name="gzyy" cols="75" rows="4"><%=rwlr_change(3)%></textarea></td>
    </tr>
	<tr>
      <td class="rtd">׼����ȡ����</td>
      <td class="ltd"><textarea name="zbfa" cols="75" rows="4"><%=rwlr_change(4)%></textarea></td>
    </tr>

	<tr>
		<td class=rtd>��ֵ</td>
		<td class=ltd><input type=text name="zf1" size=8 onblur="fzcheck();" value=<%=rs("zf")%>>��</td>
	</tr>

	<tr>
		<td class=rtd>���</td>
		<td class=ltd><input type=text name="ed" size=8 onblur="fzcheck();" value=<%=rs("ed")%>>��</td>
	</tr>

	<tr>
		<td class=rtd>�ƻ�����ʱ��</td>
		<td class=ltd>
			<select id="psy1" name="psy1" onchange='addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);'>
				<%for i = year(rs("jssj"))-1 to year(rs("jssj")) + 3%>
					<%if i = year(rs("jssj")) then%>
						<option value='<%=i%>' selected><%=i%></option>
					<%else%>
						<option value='<%= i %>'><%=i%></option>
					<%end if%>
				<%next%>
			</select>��
			<select id="psm1" name="psm1" onchange="addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);">
				<%for i = 1 to 12%>
					<%if i = month(rs("jssj")) then%>
						<option value='<%=i %>' selected><%=i%></option>
					<%else%>
						<option value='<%=i%>'><%=i%></option>
					<%end if%>
				<%next%>
			</select>��
			<select id="psd1" name="psd1">
				<%for i = 1 to 31%>
					<%if isdate(year(rs("jssj")) & "-" & month(rs("jssj")) & "-" & i) then%>
						<%if i = day(rs("jssj")) then%>
							<option value='<%=i%>' selected><%=i%></option>
						<%else%>
							<option value='<%=i%>'><%=i%></option>
						<%end if%>
					<%end if%>
				<%next%>
			</select>��
		</td>
	</tr>
  <tr>
      <td class="rtd">���η���</td>
      <td class="ltd"><input type="text" name="zrfp" size="20" value=<%=rwlr_change(5)%>></td>
    </tr>
	<tr>
		<td class=rtd>������</td>
		<td class=ltd>
			<select name="zrr1"><option></option>
				<%for i = 0 to ubound(c_jsb)%>
					<option value='<%=c_jsb(i)%>'<%if rs("zrr")=c_jsb(i) then%> selected<%end if%>><%=c_jsb(i)%></option>
				<%next%>
			</select>
		</td>
	</tr>

	<tr><td class=ctd colspan=2><input type=submit value=" �� ȷ �� �� "></td></tr></table>
	</form>
<%
end function		'ftask_change()
%>
<%
function ftaskAdd()
%>
<%Call TbTopic("�����������")%>
<form id="frm_ftask" name="ftask_add" action="ftask_indb.asp?action=add" method="post" onsubmit='return checkinf();'>
<table   class="xtable" cellspacing="0" cellpadding="3" width="80%" align="center">

    <tr>
      <th class="th" height="25">��Ŀ����
        </td>
      </th>
      <th class="th">��Ŀ����
        </td>
      </th>
    </tr>
    <tr>
      <td class="rtd" width="20%">��������</td>

      <td class="ltd"><select name="rwlx" onchange="selecttask(rwlx);">
          <option></option>
          <%for i = 0 to ubound(c_lxrwlx)%>
          <option value='<%=c_lxrwlx(i)%>'><%=c_lxrwlx(i)%></option>
          <%next%>
        </select>
	  </td>
  </tr>
</table>
<table  id="table2" class="xtable" cellspacing="0" cellpadding="3" width="80%" align="center">
    <tr>
      <td class="rtd">��������</td>
      <td class="ltd"><textarea name="rwlr" cols="75" rows="7"></textarea></td>
    </tr>
    <tr>
      <td class="rtd">��ֵ</td>
      <td class="ltd"><input type="text" name="zf" size="8" onblur="fzcheck();" />
        ��</td>
    </tr>
    <tr>
      <td class="rtd">���</td>
      <td class="ltd"><input type="text" name="ed" size="8" onblur="fzcheck();" />
        ��</td>
    </tr>
    <tr>
      <td class="rtd">����ʱ��</td>
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
    <tr>
      <td class="rtd">������</td>
      <td class="ltd"><select name="zrr">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
    </tr>

</table>
 <table  id="table1" style='display:none;'class="xtable" cellspacing="0" cellpadding="3" width="80%" align="center">
		<tr>
		<td class="rtd">������</td>
        <td class="ltd"><input type="text" name="xldh" size="30" >
         <font color="#FF0000">����</font> </td>
    </tr>
    <tr>
      <td class="rtd">�û���λ</td>
      <td class="ltd"><input type="text" name="yhdw" size="30"></td>
    </tr>
	<tr>
      <td class="rtd">ģ������</td>
      <td class="ltd"><input type="text" name="mjmc" size="30"></td>
    </tr>
	<tr>
      <td class="rtd">����С��</td>
      <td class="ltd"><input type="text" name="xlxh" size="30"></td>
    </tr>
	<tr>
		<td class="rtd">ԭ��ˮ��</td>
        <td class="ltd"><input type="text" name="ylsh" size="30">
         <font color="#FF0000">����</font> </td>
    </tr>
	<tr>
      <td class="rtd">�������������ԭ��</td>
      <td class="ltd"><textarea name="gzyy" cols="75" rows="4"></textarea></td>
    </tr>
	<tr>
      <td class="rtd">׼����ȡ����</td>
      <td class="ltd"><textarea name="zbfa" cols="75" rows="4"></textarea></td>
    </tr>
    <tr>
      <td class="rtd">��ֵ</td>
      <td class="ltd"><input type="text" name="zf1" size="8" onblur="fzcheck();" />
        ��</td>
    </tr>
    <tr>
      <td class="rtd">���</td>
      <td class="ltd"><input type="text" name="ed1" size="8" onblur="fzcheck();" />
        ��</td>
    </tr>
    <tr>
      <td class="rtd">����ʱ��</td>
      <td class="ltd"><select id="psy1" name="psy1" onchange='addOptions(this.form.psy1.value, this.form.psm1.value-1, this.form.psd1);'>
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
        <select id="psm1" name="psm1" onchange='addOptions(this.form.psy1.value, this.form.psm1.value-1, this.form.psd1);'>
          <%for i = 1 to 12%>
          <%if i = month(now) then%>
          <option value='<%=i%>' selected="selected"><%=i%></option>
          <%else%>
          <option value='<%=i%>'><%=i%></option>
          <%end if%>
          <%next%>
        </select>
        ��
        <select id="psd1" name="psd1">
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
	<tr>
      <td class="rtd">���η���</td>
      <td class="ltd"><input type="text" name="zrfp" size="20"></td>
	  </tr>
     <tr>
      <td class="rtd">������</td>
      <td class="ltd"><select name="zrr1">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
      </td>
    </tr>
</table>
<table class="xtable" cellspacing="0" cellpadding="3" width="80%" align="center">
<tr>
      <td class="ctd" colspan="2"><input type="submit" value=" �� ȷ �� �� " /></td>
    </tr>
</table>
</form>
<%
end function		'ftask_add()
%>
