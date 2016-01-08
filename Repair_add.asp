<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtask_dbinf.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->
<%
'10:29 2016-01-07
Call ChkPageAble("3,6")
CurPage="设计任务 → 添加修理任务"
strPage="mtask"
Call FileInc(0, "js/mtask.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=2 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd><%Call SearchLsh()%></td>
  </tr>
  <Tr>
    <Td class=ctd height=300><%Call mtask_add()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function mtask_add()
	Dim s_lsh, strddh, strmh, strkhmc, strdmmc, strtslb, strfz, strtszf, strmtbl, strmtjg, strdxjg, strbz, strjgzz, strsjzz, strrwxd, strjhks, strjhjg, strjhjs
	s_lsh="" : strddh="" : strmh="" : strkhmc="" : strfz="" : strtszf="" : strmtbl=40 : strmtjg="" : strdxjg="" : strbz="" 
	strdmmc="" : strjgzz="" : strsjzz="" : strrwxd=now() : strjhks=now() : strjhjg=now() : strjhjs=now()
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh<>"" Then 
		strSql="Select * from [mtask] where lsh='"&s_lsh&"'"
		Set Rs=xjweb.Exec(strSql,1)
		If not(rs.eof or rs.bof) Then
			strddh=Rs("ddh")
			strmh=Rs("mh")
			strkhmc=Rs("dwmc")
			strdmmc=Rs("dmmc")
			strtslb=Rs("tslb")
			strfz=Rs("mjzf")
			strtszf=Rs("tszf")
			strmtbl=Rs("mtbl")
			strmtjg=Rs("mtjgbl")
			strdxjg=Rs("dxjgbl")
			strbz=Rs("bz")
			strjgzz=Rs("jgzz")
			strsjzz=Rs("sjzz")
			strrwxd=Rs("rwxdsj")
			strjhks=Rs("jhkssj")
			strjhjg=Rs("jhjgsj")
			strjhjs=Rs("jhjssj")
		End If
		Rs.Close
	End If
%>
<%Call TbTopic("添改修理任务")%>
<table class=ktable cellspacing=0 cellpadding=3 width="95%">
  <form id=mtask_add name=mtask_add action=Repair_indb.asp?action=add method=post onSubmit='return checkinf();'>
    <tr>
      <th class=ctd height=25>项目名称
        </td>
      <th class=ctd>项目内容
        </td>
    </tr>
    <tr>
      <td class=rtd width="20%">修理单号</td>
      <td class=ltd><input type=text name=ddh size=30 value=<%=strddh%>></td>
    </tr>
    <tr>
      <td class=rtd>修理小号</td>
      <td class=ltd><input type=text name=lsh size=30 value=<%=s_lsh%> >
        &nbsp;<font color=red>以字母开头，且仅含数字及字母!</font></td>
    </tr>
    <tr>
      <td class=rtd>原流水号</td>
      <td class=ltd><input type=text name=mh size=30 value=<%=strmh%>></td>
    </tr>
    <tr>
      <td class=rtd>客户名称</td>
      <td class=ltd><input type=text name=dwmc size=30 value=<%=strkhmc%>></td>
    </tr>
    <tr>
      <td class=rtd>断面名称</td>
      <td class=ltd><input type=text name=dmmc size=30 value=<%=strdmmc%>>
        &nbsp;
        <select name="gfxl" onchange='changeselect(this.value);'>
          <option value="">请选择</option>
          <%for i = 0 to ubound(c_gfxl)%>
          <option value='<%=c_gfxl(i) %>'><%=c_gfxl(i)%></option>
          <%next%>
        </select>
        &nbsp;
        <select name="gfdm" onchange='this.form.dmmc.value=this.form.dmmc.value+this.value;'>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>任务分值</td>
      <td class=ltd><input type=text name=mjzf size=30 onchange=calmjfz(); value=<%=strfz%>></td>
    </tr>
    <tr>
      <td class=rtd>调试分值</td>
      <td class=ltd><input type=text name=tszf size=30 value=<%=strtszf%>></td>
    </tr>
    <tr>
      <td class=rtd>调试类别</td>
      <td colspan="2" class=ltd alt="1.选择型材类别:确定模具的最多可调试次数；<br>2.实际次数与最大次数的差说明模具结构及调试方案的正确性."><select name="tslb" style="width:51px;"  onChange="if(this.selectedIndex==0) xcts.innerHTML='';else xcts.innerHTML=' 额定调试次数:'+z_xccs[this.selectedIndex-1] + '  波动范围:' + z_xcfw[this.selectedIndex-1] + '  适用于:' + z_xcbz[this.selectedIndex-1];">
        </select>
        &nbsp;&nbsp; <span id=xcts></span></td>
    </tr>        
    <tr>
      <td class=rtd rowspan="2">分值比例</td>
      <td class=ltd>模头比例:
        <input type=text name=mtbl size=4 value=<%=strmtbl%> onchange=blchange();>
        %&nbsp;&nbsp;&nbsp;定型比例:
        <input type=text name=dxbl size=4  value=<%=100-strmtbl%> disabled>
        %</td>
    </tr>
    <tr>
        <td class=ltd>模头结构:
        <input type=text name=mtjgbl size=4 value=<%=strmtjg%>>
        %&nbsp;&nbsp;&nbsp;定型结构:
        <input type=text name=dxjgbl size=4 value=<%=strdxjg%>></td>
    </tr>
    <tr>
      <td class=rtd>备注</td>
      <td class=ltd><textarea name="bz" cols="75" rows="3" value=<%=strbz%>></textarea></td>
    </tr>       
    <tr>
      <td class=rtd>结构组长</td>
      <td class=ltd><select name="jgzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>' <%if strjgzz=c_allzz(i) then%> selected <%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>设计组长</td>
      <td class=ltd><select name="sjzz"  style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>' <%if strsjzz=c_allzz(i) then%> selected <%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>计划开始时间</td>
      <td class=ltd><script language=javascript>
  		var myDate=new dateSelector(<%=year(strjhks)&","&month(strjhks)&","&day(strjhks)%>);
  		myDate.year;
 		myDate.inputName='jhkssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>计划结构结束时间</td>
      <td class=ltd><script language=javascript>
  		var myDate=new dateSelector(<%=year(strjhjg)&","&month(strjhjg)&","&day(strjhjg)%>);
  		myDate.year;
 		myDate.inputName='jgjssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>计划全套结束时间</td>
      <td class=ltd><script language=javascript>
  		var myDate=new dateSelector(<%=year(strjhjs)&","&month(strjhjs)&","&day(strjhjs)%>);
  		myDate.year;
 		myDate.inputName='jhjssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script></td>
    </tr>
    <input type=hidden name=mjxx value="全套">
    <input type=hidden name=rwlr value="修理">    
    <input type=hidden name=bomzf value=0>
    <input type=hidden name=tsdzf value=0>
    <input type=hidden name=tsxxzlzf value=0>    
    <tr>
      <td class=ctd colspan=2><input type=submit value=" · 确 定 · "></td>
    </tr>
  </form>
</table>
<%
	call mtask_js(strtslb)
end function		'mtask_add()
%>
<%
function mtask_js(TslbOv)
'以下为JS代码%>
<script language="javascript">
//对调试类别初始化
	var z_tslb = new Array();
	var z_xcbz = new Array();
	var z_xccs = new Array();
	var z_xcfw = new Array();
<%
	set rs=xjweb.exec("select * from c_tscs order by dmlb",1)
	i=0
	do while not rs.eof
%>
		z_tslb[<%=i%>]="<%=rs("dmlb")%>";
		z_xcbz[<%=i%>]="<%=rs("bz")%>";
		z_xccs[<%=i%>]="<%=rs("edxx")%>";
		z_xcfw[<%=i%>]="<%=rs("edsx")%>";
<%
		i = i + 1
		rs.movenext
	loop
	rs.close
%>
	for(var i=1; i<z_tslb.length + 1; i++)
	{
		document.all.tslb[i] = new Option(z_tslb[i-1],z_tslb[i-1]);
		if(document.all.tslb.options[i].value=="<%=TslbOv%>")
 			document.all.tslb.options[i].selected=true; 		
	}
//模具分值初始化
	calmjfz();
	
//计算模具分值
function calmjfz()
{
	//分值系数初始化(正式使用时从库中读取)
	<%
		dim fzxs(10)
		strsql="select * from c_fzbl"
		set rs=xjweb.exec(strsql, 1)
		fzxs(0)=rs("ssgjfz")
		fzxs(1)=rs("qbfgjfz")
		fzxs(2)=rs("ryqgjfz")
		fzxs(3)=rs("ryhgjfz")
		fzxs(5)=rs("bomfzxs")
		fzxs(6)=rs("tsdfzxs")
		fzxs(7)=rs("tsfzxs")
		fzxs(8)=rs("tsxxzlfzxs")
		fzxs(9)=rs("mtjgbl")
		fzxs(10)=rs("dxjgbl")
		rs.close
	%>
	var ssgjfz=<%=fzxs(0)%>
	var qbfgjfz=<%=fzxs(1)%>
	var ryqgjfz=<%=fzxs(2)%>
	var ryhgjfz=<%=fzxs(3)%>
	var bomxs=<%=fzxs(5)%>;
	var tsdxs=<%=fzxs(6)%>;
	var tsxs=<%=fzxs(7)%>;
	var tsxxzlxs=<%=fzxs(8)%>;
	var mtjgbl=<%=fzxs(9)%>;
	var dxjgbl=<%=fzxs(10)%>;
	var str=document.all;	
	if(isNaN(parseFloat(document.all.mtjgbl.value))) document.all.mtjgbl.value=Math.round(mtjgbl*100);
	if(isNaN(parseFloat(document.all.dxjgbl.value))) document.all.dxjgbl.value=Math.round(dxjgbl*100);
	var ttmjfz=str.mjzf.value;	//模具总分

	str.bomzf.value=Math.round(ttmjfz*bomxs);
	str.tsdzf.value=Math.round(ttmjfz*tsdxs);
	str.tsxxzlzf.value=Math.round(ttmjfz*tsxxzlxs);
}

//分值比例
function blchange()
{
	var mtbl=0;
	var dxbl=0;
	if(!isNaN(parseFloat(document.all.mtbl.value))) mtbl=parseFloat(document.all.mtbl.value);
	if(!isNaN(parseFloat(document.all.dxbl.value))) dxbl=parseFloat(document.all.dxbl.value);
	if(mtbl>100) mtbl=100;
	if(dxbl>100) dxbl=100;
	if(mtbl<0) mtbl=0;
	if(dxbl<0) dxbl=0;
	dxbl=100-mtbl;
	document.all.dxbl.value=dxbl;
	document.all.mtbl.value=mtbl;
	return;
}
//端面名称关联动作
function changeselect(selvalue)
{
	var selvalue = selvalue;
	var j;
	var igfdm = new Array();
<%
	set rs=xjweb.exec("select * from c_gflb order by xl",1)
	i=0
	do while not rs.eof
%>
		igfdm[<%=i%>]=new Array("<%=rs("xl")%>","<%=rs("dm")%>");
<%
		i = i + 1
		rs.movenext
	loop
	rs.close
%>
	document.all.gfdm.length = 0;
	document.all.gfdm.options[document.all.gfdm.length] = new Option("请选择","");
	for (j=0;j<igfdm.length;j++){
		if(igfdm[j][0]==selvalue){
			document.all.gfdm.options[document.all.gfdm.length] = new Option(igfdm[j][1],igfdm[j][1]);
		}
	}
}
document.all.gfdm.options[document.all.gfdm.length] = new Option("请选择","");

</script>
<%
end function	'mtask_js()
%>