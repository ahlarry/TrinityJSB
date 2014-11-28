<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->

<%
'10:37 2011-12-07
Call ChkPageAble(3)
Call ChkDepart("技术部")
CurPage="设计任务 → 本组任务"
strPage="mtstat"
xjweb.header()
Call TopTable()

Dim iyear, imonth, dtstart, dtend, struser, irwzf, iaddfz, ilxrwzf, icount
iyear = request("searchy")
imonth = request("searchm")
struser = request("searchuser")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)

dtend=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
dtstart=cdate(iyear&"年"&imonth&"月1日")

'统计人
If struser = "" and chkable(5) Then struser = session("userName")
irwzf=0			'总分
ilxrwzf=0
iaddfz=0		'奖惩分值
icount=1		'工作项目数

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()
Sub Main()

%>

<table class="xtable" cellspacing="0" cellpadding="2" width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchMantime()%></td>
  </tr>
  <tr>
    <td class="ctd" height="280"><%Call TaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%></td>
  </tr>
</table>
<%
End Sub

Function TaskList()
	Dim iGroup, tmpSql, tmpRs, ArrZrr, ArrJs, ArrFz, ArrSjfz, n
	Call TbTopic(struser & "组" & imonth & "月任务定额")
%>
<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable"  align="center">
  <tr>
    <th class="th" width="20">id
      </td>
    </th>
    <th class="th">流水号
      </td>
    <th class="th" width="120">任务内容
      </td>
    </th>
    <th class="th">责任人
      </td>
    </th>
    <th class="th" width="120">角色
      </td>
    </th>
    <th class="th" width="*">定额断面
      </td>
    </th>
    <th class="th" width="100">常规定额
      </td>
    </th>
    <th class="th" width="100">本套定额
      </td>
    </th>
  </tr>
  <%
	tmpSql="Select [user_group] from [ims_user] where [user_name]='"&struser&"'"
		Set tmpRs=xjweb.Exec(tmpSql,1)
		If Not(tmpRs.Eof Or tmpRs.Bof) Then
			iGroup=tmpRs("user_group")
		Else
			iGroup=0
		End If
	tmpRs.Close
				  
  	strSql="select a.*, b.* ,a.lsh as lsh, a.xz as xz,b.rwlr as rwlr from [reward] a, [mtask] b where xz="&iGroup&" and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 and a.lsh=b.lsh order by jssj desc, a.lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Dim tmprw, tmplsh, tmpdm, tmpfz, tmpbz, dtlsh 		'单套流水号分值
	tmprw="" : tmplsh="" : tmpdm="" : tmpbz=1 : tmpfz=0 : dtlsh=0
	Do While Not Rs.eof
		If tmplsh<>"" and tmplsh<>Rs("lsh") Then
			ArrZrr=split(ArrZrr,",")
			ArrJs=split(ArrJs,",")
			ArrFz=split(ArrFz,",")
			ArrSjfz=split(ArrSjfz,",")
	%>
  <tr onclick="show(<%=icount%>)">
    <td class=ctd><img  id=<%="img"&icount%> src="images/plus.png" width="16" height="16" alt="展开" /><%=icount%></td>
    <td class=ctd><a href="mtask_display.asp?s_lsh=<%=tmplsh%>"><%=tmplsh%></a></td>
    <td class=ctd><%=tmprw%></td>
    <td class=ctd>&nbsp;</td>
    <td class=ctd>&nbsp;</td>
    <td class=ctd><%=tmpdm%></td>
    <td class=ctd><%=tmpfz%></td>
    <td class=ctd title=<%=tmpbz%>><%=dtlsh%></td>
  </tr>
<tbody id="child<%=icount%>" style="display:none;" >  
  <%for n=0 to ubound(ArrZrr)%>
  <tr>
    <td class=rtd colspan="4"><%=ArrZrr(n)%></td>
    <td class=ctd colspan="2"><%=ArrJs(n)%></td>
    <td class=ctd><%=Round(ArrFz(n),1)%></td>
    <td class=ctd><%=Round(ArrSjfz(n),1)%></td>
  </tr>
  <%next%>
  </tbody>
  <%
			dtlsh=0 : ArrZrr="" : ArrJs="" : ArrFz="" : ArrSjfz="" : tmpfz=0
			icount = icount + 1
		Else
			If ArrZrr="" Then ArrZrr=Rs("zrr") else ArrZrr=ArrZrr & "," & Rs("zrr")
			If ArrJs="" Then ArrJs=Rs("js") else ArrJs=ArrJs & "," & Rs("js")
			If ArrFz="" Then ArrFz=Rs("fz") else ArrFz=ArrFz & "," & Rs("fz")
			If ArrSjfz="" Then ArrSjfz=Rs("sjfz") else ArrSjfz=ArrSjfz & "," & Rs("sjfz")
			tmpbz=Rs("a.bz")
			tmpfz=tmpfz + Rs("fz")
			dtlsh=dtlsh + Rs("sjfz")
		End If
		tmplsh=Rs("lsh")
		tmpdm=Rs("dedm")
		If Rs("demt")>0 Then tmprw="模头"&Rs("mtrw")
		If Rs("dedx")>0 Then tmprw=tmprw&" 定型"&Rs("dxrw")
		irwzf=irwzf+Round(Rs("sjfz"),1)
		Rs.movenext
	loop
	%>
  <tr>
    <td class=rtd colspan=7>设计任务总分:</td>
    <td class=ctd><b><%=Round(irwzf,1)%></b></td>
  </tr>
  <%
	Rs.close
	Dim mystr, mystr1, rwlr_change
	strSql="select * from [ftask] where xz="&iGroup&" and (rwlx='零星修理' or rwlx='零星任务' or rwlx='技术代表设计') and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by rwlx desc,jssj desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Do While Not Rs.eof
		mystr=rs("rwlr")
	    mystr=split(mystr,"||")
		If 5 > ubound(mystr) Then
		  	mystr1=split(rs("rwlr"),":")
		   	rwlr_change=""
		  else
	   	 	mystr1=mystr(5)
			mystr1=split(mystr1,":")
			rwlr_change=mystr1(1)
	   End If
	   If not(Rs("rwlx")="零星修理" and InStr(rwlr_change, "设计")=0) Then
%>
  <tr>
    <td class=ctd><%=icount%></td>
    <td class=ctd><%If Rs("rwlx")="技术代表设计" Then Response.Write(mystr1(0)) else Response.Write(Rs("xldh")) End If%>
      &nbsp;</td>
    <td class=ctd><%=Rs("rwlx")%></td>
    <td class=ctd><%=Rs("zrr")%></td>
    <td class=ctd><%=rwlr_change%>&nbsp;</td>
    <td class=ctd>&nbsp;</td>
    <td class=ctd><%=Rs("ed")%></td>
    <td class=ctd><%=Rs("ed")%></td>
  </tr>
  <%
		End If
		icount = icount + 1
		ilxrwzf=ilxrwzf+Rs("ed")
		Rs.movenext
	loop
%>
  <tr>
    <td class=rtd colspan=7>零星任务总分:</td>
    <td class=ctd><b><%=ilxrwzf%></b></td>
  </tr>
  <%
	Rs.close	
  %>
  <tr>
    <td class=rtd colspan=7>任务总分:</td>
    <td class=ctd><b><%=Round(irwzf+ilxrwzf,1)%></b></td>
  </tr>
</table>
<%
End Function

Function SearchMantime()
%>
<table cellpadding=2 cellspacing=0>
  <form action=<%=request.servervariables("script_name")%> method=get>
    <tr>
      <td> 请选择:
        <select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <%for i = year(now) - 3 to year(now)%>
          <option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <%for i = 1 to 12%>
          <option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
          <%next%>
        </select>
        月&nbsp;
        <select name="searchuser" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value+"&searchuser="+this.form.searchuser.value);'>
          <option value=""></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value="<%=c_allzz(i)%>" <%If struser = c_allzz(i) Then%>selected<%end If%>><%=c_allzz(i)%></option>
          <%next%>
        </select>
        <input type="submit" value=" 选 择 "></td>
    </tr>
  </form>
</table>
<%
End Function
%>
<script language="javascript">
function pucker_show(name,no,hiddenclassname,showclassname) {
    //name:命名前缀
    //no:对象的序号
    //showclassname:展开状态样式名
    //hiddenclassname:折叠状态样式名
    for (var i=1 ;i<6 ;i++ )
    {
        document.getElementById(name+i).className=hiddenclassname;
    }
    document.getElementById(name+no).className=showclassname;
}

function show(num){
	var obj1=document.getElementById("img"+num) 
	if(obj1.src.indexOf("images/minus.png")>0){
		obj1.src="images/plus.png";
		}
	else
		{
		obj1.src="images/minus.png";
		}
			
	var obj2=document.getElementById("child"+num) 
	obj2.style.display=(obj2.style.display=="")?"none":"" 
}	
</script> 
