<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/ftask_dbinf.asp"-->
<%
Call ChkPageAble(3)
CurPage="零星任务 → 更改零星任务"
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
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%Call ftaskChange()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function ftaskChange()
	Dim iid
	iid=0
	iid=Trim(Request("id"))
	If iid=0 Or iid="" Or Not(IsNumeric(iid)) Then Call JsAlert("请从正规的入口进入!","ftask_list.asp") : Exit Function

	strSql="select * from [ftask] where id="&iid&""
	Set Rs=xjweb.Exec(strSql,1)
	if rs.eof or rs.bof then
		Call JsAlert("ID号 〖"&iid&"〗的零星任务不存在!","")
	else
		call ftask_change(rs)
	end if
	rs.close
End Function
Function rwlr_change(i)
         dim mystr,mystr1,mystr2
	     mystr2=rs("rwlx")
		 mystr=rs("rwlr")
		 if mystr2="零星修理" then
			 If Instr(mystr,"||")>0 Then
			     mystr=split(mystr,"||")
			     If i > ubound(mystr) Then
			     	mystr1=""
			     	rwlr_change=mystr1
			     else
	   		 		 mystr1=mystr(i)
					 mystr1=split(mystr1,":")
					 rwlr_change=mystr1(1)
				 End If

			 else
			    mystr=split(mystr,chr(10))
	   		 	mystr1=mystr(i)
	   		 	If Instr(mystr1,"：")>0 Then
					mystr1=split(mystr1,"：")
					rwlr_change=mystr1(1)
				else
					rwlr_change=mystr1
				End If
			 End If
		 else
		 rwlr_change=mystr
		 end if
End Function


function ftask_change(rs)

%>
	<%Call TbTopic("更改零星任务")%>
	<%call rwlr_change(i)%>
	<form id=frm_ftask name=frm_ftask action=ftask_indb.asp?action=change method=post onSubmit='return checkinf();'>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">

	<tr>
		<th class=th height=25>项目名称</th>
		<th class=th>项目内容</th>
	</tr>
	<tr>
		<td class=rtd width="20%">任务类型</td>
		<td class=ltd>
			<select name="rwlx"><option value=<%=rs("rwlx")%>><%=rs("rwlx")%></option>
				<%for i = 0 to ubound(c_lxrwlx)%>
					<option value='<%=c_lxrwlx(i)%>'><%=c_lxrwlx(i)%></option>
				<%next%>
			</select>
		</td>
	</tr>
	</table>
	<%if rs("rwlx")="零星修理" then %>

	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<tr>
		<td class="rtd">修理单号</td>
        <td class="ltd"><input type="text" name="xldh" size="30"  value=<%=rs("xldh")%>>
         <font color="#FF0000">必填</font> </td>
    </tr>
    <tr>
      <td class="rtd">用户单位</td>
      <td class="ltd"><input type="text" name="yhdw" size="30" value=<%=rwlr_change(0)%>></td>
    </tr>
	<tr>
      <td class="rtd">模具名称</td>
      <td class="ltd"><input type="text" name="mjmc" size="30" value=<%=rwlr_change(1)%>></td>
    </tr>
	<tr>
      <td class="rtd">修理小号</td>
      <td class="ltd"><input type="text" name="xlxh" size="30" value=<%=rwlr_change(2)%>></td>
    </tr>
	<tr>
		<td class="rtd">原流水号</td>
        <td class="ltd"><input type="text" name="ylsh" size="30"  value=<%=rwlr_change(6)%>>
         <font color="#FF0000">必填</font> </td>
    </tr>
	<tr>
      <td class="rtd">故障现象与分析原因</td>
      <td class="ltd"><textarea name="gzyy" cols="75" rows="4"><%=rwlr_change(3)%></textarea></td>
    </tr>
	<tr>
      <td class="rtd">准备采取方案</td>
      <td class="ltd"><textarea name="zbfa" cols="75" rows="4"><%=rwlr_change(4)%></textarea></td>
    </tr>

	<tr>
		<td class=rtd>分值</td>
		<td class=ltd><input type=text name="zf1" size=8 onblur="fzcheck();" value=<%=rs("zf")%>>分</td>
	</tr>
	
	<tr>
		<td class=rtd>额度</td>
		<td class=ltd><input type=text name="ed" size=8 onblur="fzcheck();" value=<%=rs("ed")%>>分</td>
	</tr>

	<tr>
		<td class=rtd>计划结束时间</td>
		<td class=ltd>
			<select id="psy1" name="psy1" onchange='addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);'>
				<%for i = year(rs("jssj"))-1 to year(rs("jssj")) + 3%>
					<%if i = year(rs("jssj")) then%>
						<option value='<%=i%>' selected><%=i%></option>
					<%else%>
						<option value='<%= i %>'><%=i%></option>
					<%end if%>
				<%next%>
			</select>年
			<select id="psm1" name="psm1" onchange="addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);">
				<%for i = 1 to 12%>
					<%if i = month(rs("jssj")) then%>
						<option value='<%=i %>' selected><%=i%></option>
					<%else%>
						<option value='<%=i%>'><%=i%></option>
					<%end if%>
				<%next%>
			</select>月
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
			</select>日
		</td>
	</tr>
  <tr>
      <td class="rtd">责任分配</td>
      <td class="ltd"><input type="text" name="zrfp" size="20" value=<%=rwlr_change(5)%>></td>
    </tr>
	<tr>
		<td class=rtd>责任人</td>
		<td class=ltd>
			<select name="zrr1"><option value='<%=rs("zrr")%>'><%=rs("zrr")%></option>
				<%for i = 0 to ubound(c_jsb)%>
					<option value='<%=c_jsb(i)%>'<%if rs("zrr")=c_jsb(i) then%> selected<%end if%>><%=c_jsb(i)%></option>
				<%next%>
			</select>
		</td>
	</tr>
	</table>
	<%else%>

	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<tr>
		<td class=rtd>任务内容</td>
		<td class=ltd><textarea name="rwlr" cols="75" rows="7"><%=rwlr_change(0)%></textarea></td>
	</tr>

	<tr>
		<td class=rtd>分值</td>
		<td class=ltd><input type=text name="zf" size=8 onblur="fzcheck();" value=<%=rs("zf")%>>分</td>
	</tr>

	<tr>
		<td class=rtd>额度</td>
		<td class=ltd><input type=text name="ed" size=8 onblur="fzcheck();" value=<%=rs("ed")%>>分</td>
	</tr>

	<tr>
		<td class=rtd>计划结束时间</td>
		<td class=ltd>
			<select id="psy" name="psy" onchange='addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);'>
				<%for i = year(rs("jssj"))-1 to year(rs("jssj")) + 3%>
					<%if i = year(rs("jssj")) then%>
						<option value='<%=i%>' selected><%=i%></option>
					<%else%>
						<option value='<%= i %>'><%=i%></option>
					<%end if%>
				<%next%>
			</select>年
			<select id="psm" name="psm" onchange="addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);">
				<%for i = 1 to 12%>
					<%if i = month(rs("jssj")) then%>
						<option value='<%=i %>' selected><%=i%></option>
					<%else%>
						<option value='<%=i%>'><%=i%></option>
					<%end if%>
				<%next%>
			</select>月
			<select id="psd" name="psd">
				<%for i = 1 to 31%>
					<%if isdate(year(rs("jssj")) & "-" & month(rs("jssj")) & "-" & i) then%>
						<%if i = day(rs("jssj")) then%>
							<option value='<%=i%>' selected><%=i%></option>
						<%else%>
							<option value='<%=i%>'><%=i%></option>
						<%end if%>
					<%end if%>
				<%next%>
			</select>日
		</td>
	</tr>

	<tr>
		<td class=rtd>责任人</td>
		<td class=ltd>
			<select name="zrr"><option value='<%=rs("zrr")%>'><%=rs("zrr")%></option>
				<%for i = 0 to ubound(c_jsb)%>
					<option value='<%=c_jsb(i)%>'<%if rs("zrr")=c_jsb(i) then%> selected<%end if%>><%=c_jsb(i)%></option>
				<%next%>
			</select>
		</td>
	</tr>
	</table>
	<%end if%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<tr><td class=ctd colspan=2><input type=submit value=" · 确 定 · "></td></tr>
	<input type="hidden" name=id value=<%=rs("id")%>>
	</form>
	</table>
<%
end function		'ftask_change()
%>
