<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
Call ChkPageAble("3,4")
CurPage="调试任务 → 齐套信息整理任务分配"				
strPage="atask"
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Dim strterm
strterm = trim(request("lsh"))

Sub Main()
%>
<table class="xtable" cellspacing="0" cellpadding="0" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd"><%Call InfoFixAssign()%>
      <%Response.Write(XjLine(1,"100%",web_info(12)))%>
      <%Call InfoFix_nofinished() %>
      <%Response.Write(XjLine(10,"100%",""))%>
    </td>
  </tr>
</table>
<%
End Sub

Function InfoFixAssign()%>
<%Call TbTopic("添加模具齐套信息整理任务书")%>
<table id="table1" class="ktable" cellspacing="0" cellpadding="3" width="98%">
  <tr>
    <th class="rtd" width="15%">项目名称
      </td>
    </th>
    <th class="ltd">项目内容
      </td>
    </th>
  </tr>
  <tr>
    <td class="rtd">流水号:</td>
    <td class="ltd"><input id="lsh" name="lsh" type="text" onchange='FindLsh();' size="15" />
      模具总分:<span id=span_mjzf style="font-weight:bold;">0</span>分
      <input type=hidden name=mjzf value=0>
    </td>
  </tr>
  <tr>
    <td class="rtd">任务系数:</td>
    <td class="ltd"><input name="rwxs" type="text" value="0.08" size="8" />
    </td>
  </tr>
  <tr>
    <td class="rtd">共同完成人数:</td>
    <td class="ltd"><input name="zxrs" type="text" id="zxrs" value="1" onchange="tdvalue(this.value);" size="8"/>
    <input id="btnDelete" name="btnDelete" type="button" onclick="DeleteTableRow()" value="减少" />
    </td>
  </tr>
  <tr>
    <td class="rtd">任务人:</td>
    <td class="ltd"><select name="rwr0">
        <option></option>
        <%for i = 0 to ubound(c_allzy)%>
        <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
        <%next%>
      </select>
      &nbsp;&nbsp;&nbsp;&nbsp;分值系数:
      <input name="xs0" type="text" onchange="fzhi()" value="1" size="5" />
    </td>
  </tr>
  <tr>
    <td class="rtd">计划开始时间</td>
    <td class="ltd"><select id="jhksy" name="jhksy" onchange='addOptions(this.form.jhksy.value, this.form.jhksm.value-1, this.form.jhksd);'>
        <%for i = year(now)-1 to year(now) + 3%>
        <%if i = year(now) then%>
        <option value='<%=i%>' selected="selected"><%=i%></option>
        <%else%>
        <option value='<%=i%>'><%=i%></option>
        "
        <%end if%>
        <%next%>
      </select>
      年
      <select id="jhksm" name="jhksm" onchange='addOptions(this.form.jhksy.value, this.form.jhksm.value-1, this.form.jhksd);'>
        <%for i = 1 to 12%>
        <%if i = month(now) then%>
        <option value='<%=i%>' selected="selected"><%=i%></option>
        <%else%>
        <option value='<%=i%>'><%=i%></option>
        <%end if%>
        <%next%>
      </select>
      月
      <select id="jhksd" name="jhksd">
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
      日 </td>
  </tr>
  <tr>
    <td class="rtd">计划结束时间</td>
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
      年
      <select id="psm" name="psm" onchange='addOptions(this.form.psy.value, this.form.psm.value-1, this.form.psd);'>
        <%for i = 1 to 12%>
        <%if i = month(now) then%>
        <option value='<%=i%>' selected="selected"><%=i%></option>
        <%else%>
        <option value='<%=i%>'><%=i%></option>
        <%end if%>
        <%next%>
      </select>
      月
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
      日 </td>
  </tr>
  <tr>
    <td class="rtd">备注:</td>
    <td valign="middle" class="ltd"><textarea name="bz" cols="65" rows="6"></textarea>
      <input name="submit" type="submit" value=" · 确 定 · " /></td>
  </tr>
</table>
<%
End Function

Function InfoFix_nofinished()			'具有分配权限的未完成的调试任务
	Dim RecordPerPage,absPageNum,absRecordNum,iCounter,TotalCount
	absPageNum = 0
	RecordPerPage = 10
	iCounter = 1
	strSql="select * from [mtask] where not(mjjs) and ((mjxx='全套'and not(isnull(mttsxxzlr)) and not(isnull(dxtsxxzlr))) or (mjxx='模头'and not(isnull(mttsxxzlr))) or (mjxx='定型'and not(isnull(dxtsxxzlr)))) and isnull(xtxxzljs) order by lsh desc"
	Call xjweb.Exec("",-1)
	Set Rs=Server.CreateObject("ADODB.RECORDSET")
	Rs.CacheSize=RecordPerPage
	Rs.open strSql,Conn,1,3
	If (Rs.Eof Or Rs.Bof) Then
		Call TbTopic("暂时没有待分配辅助任务!") : Exit Function
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
<%Call TbTopic("共有 "& rs.recordcount &" 套待分配的辅助任务")%>
<table width="95%" cellspacing="0" cellpadding="2" class="xtable">
  <tr>
    <th class="th">id</th>
    <th class="th">订单号</th>
    <th class="th">流水号</th>
    <th class="th">单位名称</th>
    <th class="th">断面名称</th>
    <th class="th">技术代表</th>
    <th class="th">任务内容</th>
    <th class="th" width="*">调试结束时间</th>
    <th class="th" width="*">调试信息整理</th>
    <th class="th" width="*">模具齐套信息整理</th>
  </tr>
  <% for absrecordnum = 1 to recordperpage %>
  <tr>
    <td class="ctd"><%=icounter %></td>
    <td class="ctd"><%= rs("ddh")%></td>
    <td class="ctd"><a href="InfoFix_assign.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="ctd"><%=rs("dwmc")%></td>
    <td class="ctd"><%=rs("dmmc")%></td>
    <td class="ctd"><%=rs("jsdb")%></td>
    <td class="ctd"><%= rs("mjxx") & rs("rwlr") %></td>
    <td class="ctd"><%If rs("mttsjs")<>"" Then Response.Write(FormatDateTime(rs("mttsjs"),2)) else Response.Write(FormatDateTime(rs("dxtsjs"),2)) End If%></td>
    <%select case rs("mjxx")%>
    <%case "全套"%>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
      <%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxzljs"),rs)%>
    </td>
    <%case "模头"%>
    <td class="ctd"><%call distd(rs("mttsdks"),rs("mttsdjs"),-20,rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsks"),rs("mttsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxzljs"),rs)%>
    </td>
    <%case "定型"%>
    <td class="ctd"><%call distd(rs("dxtsdks"),rs("dxtsdjs"),-20,rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("dxtsks"),rs("dxtsjs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxzljs"),rs)%>
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
<table width="95%" cellpadding="2" cellspacing="0" border="0">
  <tr>
    <td align="left"> 符合条件共 <%=rs.recordcount%> 个&nbsp;&nbsp;
      每页 <%=rs.pagesize%> 个&nbsp;&nbsp;
      共 <%=Rs.PageCount%> 页&nbsp;&nbsp;
      当前为第 <%=absPageNum%> 页 </td>
    <td align="right"> 【
      <%
				if absPageNum > 1 then
					response.write("<a href="&Request.ServerVariables("script_name")&"?ipage="&(abspagenum-1)&" alt='上一页'> ←</a>&nbsp;&nbsp;")
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
					response.write("&nbsp;<a href="&Request.ServerVariables("script_name")&"?ipage="&(absPageNum+1)&" alt='下一页'> → </a>&nbsp;")
				end if
			%>
      】
      跳转到:
      <select name="ipage" onchange='location.href(&quot;<%=Request.ServerVariables("script_name")%>?ipage=&quot; + this.value+&quot;&quot;);'>
        <%for i=1 to Rs.PageCount%>
        <%if i = absPageNum then%>
        <option value="<%=i%>" selected="selected">第 <%=i%> 页</option>
        <%else%>
        <option value="<%=i%>">第 <%=i%> 页</option>
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
<script language="JavaScript" type="text/javascript">
		var objdoc=document.all;
		function checkjs()
		{
			var str;
			str=objdoc.form_assign.fplr.value
			if (str.substr(0,2)=='结束')
				objdoc.form_assign.zrr.disabled=true;
			else
				objdoc.form_assign.zrr.disabled=false;
		}
		checkjs();

		function checksubinf(frm)
		{
			if (frm.fplr.value==""){alert("请选择分配内容!"); frm.fplr.focus(); return false;}
			if ((frm.zrr.value=="") && (!frm.zrr.disabled)){alert("请选择责任人!"); frm.zrr.focus(); return false;}
			return true;
		}
	</script>
<%
end function
%>
<script language="javascript">
function FindLsh()
{
document.all.lsh.style.color='green';
	var ttmjfz=0;		//模具总分值
	var str=document.all;
	//由参考断面获得初始分值
	ttmjfz=str.lsh.value
	str.span_mjzf.innerHTML=ttmjfz;
	str.mjzf.value=ttmjfz;
}

function fzhi()
{
}

 var   objTable;
 function   PageLoad()
 {
         objTable = document.getElementById("table1");                 //找到操作Table
 }
 function tdvalue(temprs)
 {
         var objTempRow =objTable.rows[4];         //找到Table的模版行
 for   (var n=1; i<temprs; i++)   
 {  
    var objNewRow=objTable.insertRow(5);             //新增一行
 	objNewRow.id   =   objTable.rows.length   -   1;
 	//以模版行建立新行内容
 	for   (var i=0; i<objTempRow.cells.length; i++)
 	{
      	   var objNewCell=objNewRow.insertCell(i);
       	  objNewCell.innerHTML = objTempRow.cells[i].innerHTML;
	 }
 }
}
function DeleteTableRow()
{
        if   (objTable.rows.length-1>7)
        {
                objTable.deleteRow(5);
        }
}
 window.onload   =   PageLoad;
</script>
