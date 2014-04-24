<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtask_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
'10:52 2007-1-25-星期四
Call ChkPageAble("3,4")
CurPage="调试任务 → 齐套信息整理任务分配"
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
            <td>&nbsp;&nbsp;输入订单号:
              <input tabindex="1" type="text" name="s_ddh" size="15" value="<%=Trim(Request("s_ddh"))%>" />
            </td>
            <td>&nbsp;&nbsp;输入流水号:
              <input tabindex="1" type="text" name="s_ls" size="15" value="<%=Trim(Request("s_ls"))%>" />
              <input type="submit" value=" 查 找 " />
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
排序:
<select name="order" onchange='location.href(&quot;<%=request.servervariables("script_name")%>?ipage=1&amp;order=&quot; + this.value);'>
  <option value="ddh" selected="selected">订单号</option>
  <option value="lsh" <%If strOrder="lsh" Then%>selected<%End If%>>流水号</option>
  <option value="khmc" <%If strOrder="khmc" Then%>selected<%End If%>>客户名称</option>
</select>
<%
End Function

Function InfoFix_add()
	Dim s_ddh, s_ls, s_time
	s_ddh="" : s_ls="" : s_time=""
	If Trim(Request("s_ddh"))<>"" Then s_ddh=Trim(Request("s_ddh"))
	If Trim(Request("s_ls"))<>"" Then s_ls=Trim(Request("s_ls"))
	If ( s_ddh="" and s_ls="") Then Call TbTopic("请输入要更改的任务书的订单号或流水号!") : Exit Function

'	strSql="select a.*, b.tscs, b.tscs as tscs from [ts_mould] b, [mtask] a where [ddh]='"&s_ddh&"' and ((mjxx='全套'and not(isnull(mttsxxzlr)) and not(isnull(dxtsxxzlr))) or (mjxx='模头'and not(isnull(mttsxxzlr))) or (mjxx='定型'and not(isnull(dxtsxxzlr)))) and isnull(xtxxsjjs) and b.tscs>1 and a.lsh=b.lsh"
	strSql="select * from [mtask] a where ([ddh]='"&s_ddh&"' or [lsh]='"&s_ls&"') and isnull(xtxxsjjs) and (not(isNull(mttsdjs)) or not(isNull(dxtsdjs)))"
	set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		If s_ddh<>"" Then
			Call JsAlert("订单号为 【 " & s_ddh & " 】 的任务书不存在或已结束!", "InfoFix_add.asp")
		Else
			If s_ls<>"" Then
				Call JsAlert("流水号为 【 " & s_ls & " 】 的任务书不存在或已结束!", "InfoFix_add.asp")
			End If
		End If
	Else
			Call InfoFixadd(Rs)
	End If
	Rs.Close
End Function

Function InfoFixadd(Rs)
Call TbTopic("添加模具齐套信息整理任务书")
%>
<table id="table1" class="ktable" cellspacing="0" cellpadding="3" width="98%" align="center">
  <form id="InfoFix_add" name="InfoFix_add" action="InfoFix_indb.asp?action=add" method="post">
    <tr>
      <th class="rtd" width="15%">项目名称
        </td>
      </th>
      <th class="ltd">项目内容
        </td>
      </th>
    </tr>
    <tr>
      <td class="rtd">订单号:</td>
      <td class="ltd"><input name="lsh" type="text" disabled="disabled" id="ddh" onchange="FindLsh();" value="<%=Rs("ddh")%>" size="15" />
        &nbsp;&nbsp;&nbsp;&nbsp;
        任务总分:<span id="span_rwzf" style="font-weight:bold;">0</span>分 </td>
    </tr>
    <%
Dim m,sjxs
m=1 : sjxs=0
while not rs.eof
	If Rs("xtxxzlr")<>"" Then
%>
    <tr>
      <td class="rtd">执行人:</td>
      <td class="ltd"><%=Rs("xtxxzlr")%> &nbsp;&nbsp;流水号:<%=Rs("lsh")%>&nbsp;&nbsp;分值系数:<%=Rs("xtxxzlxs")%> &nbsp;&nbsp;任务分:<%=Round(Rs("mjzf")*Rs("xtxxzlxs"),1)%>&nbsp;&nbsp;计划结束时间:<%=Rs("xtxxjhjs")%>&nbsp;&nbsp;
        <select name="xtzlwc<%=m%>">
          <option></option>
          <option value='整理结束'>齐套信息整理结束</option>
          <input type="hidden" name="zxr<%=m%>" value="<%=Rs("xtxxzlr")%>" />
          <input type="hidden" name="lsh<%=m%>" value="<%=Rs("lsh")%>" />
          <input type="hidden" name="fzxs<%=m%>" value="" />
          <input type="hidden" name="zlxf<%=m%>" value=<%=Round(Rs("mjzf")*Rs("xtxxzlxs"),1)%> />
        </select>
      </td>
    </tr>
    <%else%>
    <tr>
      <td class="rtd">执行人:</td>
      <td class="ltd"><select name="zxr<%=m%>" onchange="fzhi(<%=m%>,1);">
          <option></option>
          <%for i = 0 to ubound(c_jsb)%>
          <option value='<%=c_jsb(i)%>'><%=c_jsb(i)%></option>
          <%next%>
        </select>
        &nbsp;&nbsp;&nbsp;&nbsp;流水号:<%=Rs("lsh")%> &nbsp;&nbsp;模具分:<%=Rs("mjzf")%>&nbsp;&nbsp;&nbsp;&nbsp;分值系数:
        <input name="fzxs<%=m%>" type="text" onchange="fzhi(<%=m%>,2);" value="0.08" size="5" />
        &nbsp;&nbsp;任务分:<span id="span_zldf<%=m%>" style="font-weight:bold;">0</span>分
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
    <%End If%>
    <tr>
      <td colspan="2" align="center"><input name="submit" type="submit" value=" · 确 定 · " />
      </td>
    </tr>
  </form>
</table>
<%
End Function	'End InfoFixadd()

Function InfoFix_nofinished()			'具有分配权限的未完成的调试任务
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

'	strSql="select a.*, b.tscs, b.tscs as tscs from [ts_mould] b, [mtask] a where not(mjjs) and ((mjxx='全套'and not(isnull(mttsxxzlr)) and not(isnull(dxtsxxzlr))) or (mjxx='模头'and not(isnull(mttsxxzlr))) or (mjxx='定型'and not(isnull(dxtsxxzlr)))) and isnull(xtxxsjjs) and b.tscs>1 and a.lsh=b.lsh order by a.ddh desc"
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
<table width="95%" cellspacing="0" cellpadding="2" class="xtable" align="center">
  <tr>
    <th class="th">id</th>
    <th class="th">订单号</th>
    <th class="th">流水号</th>
    <th class="th">单位名称</th>
    <th class="th">断面名称</th>
    <th class="th">技术代表</th>
    <th class="th" width="*">调试信息整理</th>
    <th class="th" width="*">模具齐套信息整理</th>
    <th class="th" width="*">齐套整理计划结束</th>
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
    <%case "全套"%>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%>
      <%call distd2(rs("dxtsxxzlks"),rs("dxtsxxzljs"),rs)%>
    </td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%>
    </td>
    <td class="ctd"><%=Rs("xtxxjhjs")%>&nbsp;</td>
    <%case "模头"%>
    <td class="ctd"><%call distd2(rs("mttsxxzlks"),rs("mttsxxzljs"),rs)%></td>
    <td class="ctd"><%call distd2(rs("xtxxzlks"),rs("xtxxsjjs"),rs)%></td>
    <td class="ctd"><%=Rs("xtxxjhjs")%>&nbsp;</td>
    <%case "定型"%>
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
%>
<script language="JavaScript" type="text/javascript">
 function fzhi(x,y)		//根据任务系数和共同完成人数来确定个人系数及分值
{
	var tmpfz=0;
	eval("document.all.span_zldf" + x + ".innerHTML=Math.round(document.all.mjf" + x + ".value*document.all.fzxs" + x + ".value*100)/100.0;");	//显示本流水号分值
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
