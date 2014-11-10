<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtask_dbinf.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->
<%
'10:52 2007-1-25-星期四
Call ChkPageAble(3)
CurPage="设计任务 → 更改任务书"
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
    <Td class=ctd height=300><%Call mtaskChange()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function mtaskChange()
	Dim s_lsh, s_time, CC
	s_lsh="" : s_time="" : CC=0
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("请输入要更改的任务书的流水号!") : Exit Function

	DO while CC<2
	strSql="select * from [mtask] where [lsh]='"&s_lsh&"'"
	set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		s_lsh="C"&s_lsh
		CC=CC+1
		If CC=2 Then
			Call JsAlert("流水号为 【 " & right(s_lsh,len(s_lsh)-1) & " 】 的任务书不存在!", "mtask_change.asp")
		End If
	Else
		CC=3
		If Not(IsNull(Rs("sjjssj"))) Then
			s_time=datediff("d", rs("sjjssj"), now)
			if s_time > 5 then
	'			Call JsAlert("流水号为 【 " & s_lsh & " 】 的任务书已经完成"& s_time &" 天,已过可编辑期限!", "mtask_change.asp")
				Call BzChan(Rs)
			else %>
<Script language="javascript">
				alert("【<%=s_lsh%>】任务书已完成,<%=5-s_time%>天之内仍可更改。");
				</Script>
<% Call mtask_change(Rs)
			End if
		Else
			Call mtask_change(Rs)
		End if
	End If
	Loop
	'Rs.Close
End Function

Function mtask_change(Rs)
%>
<%Call TbTopic("更改流水号 <font style=color:#0000FF>" & Rs("lsh") & "</font> 的任务书")%>
<%If ChkAble(4) Then Response.Write "<a href=mtask_zzchange.asp?s_lsh="&Rs("lsh")&">组长权限</a><br>"%>
<table class=xtable cellspacing=0 cellpadding=3 width="95%">
  <form id=mtask_add name=mtask_add action=mtask_indb.asp?action=change method=post onSubmit='return checkinf();'>
    <tr>
      <th class=th height=25>项目名称
        </td>
      <th class=th>项目内容
        </td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>■合同信息■</b></td>
      <td class=ctd>&nbsp;</td>
    </tr>
    <tr>
      <td class=rtd width="20%">订单号</td>
      <td class=ltd><input type=text name=ddh size=30 value=<%=Rs("ddh")%>></td>
    </tr>
    <tr>
      <td class=rtd>流水号</td>
      <td class=ltd><input type=text name=lsh size=30 value=<%=Rs("lsh")%>></td>
    </tr>
    <tr>
      <td class=rtd>模号</td>
      <td class=ltd><input type=text name=mh size=30 value=<%=Rs("mh")%>></td>
    </tr>
    <tr>
      <td class=rtd>断面名称</td>
      <td class=ltd><input type=text name=dmmc size=30 value="<%=Rs("dmmc")%>">
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
      <td class=rtd>客户名称</td>
      <td class=ltd><input type=text name=dwmc size=30 value=<%=Rs("dwmc")%>>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onChange='this.form.dwmc.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_dwmc)%>
          <option value='<%=c_dwmc(i)%>'><%=c_dwmc(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>设备厂家</td>
      <td class=ltd><input type=text name=sbcj size=30 value=<%=Rs("sbcj")%>>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onchange='this.form.sbcj.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_sbcj)%>
          <option value='<%=c_sbcj(i) %>'><%=c_sbcj(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>挤出机型号</td>
      <td class=ltd><input type=text name=jcjxh size=30 value=<%=Rs("jcjxh")%>>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onchange='this.form.jcjxh.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_jcjxh)%>
          <option value='<%=c_jcjxh(i)%>'><%=c_jcjxh(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>模具材料</td>
      <td class=ltd><input type=text name=mjcl size=30 value=<%=Rs("mjcl")%>>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onchange='this.form.mjcl.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_mjcl)%>
          <option value='<%=c_mjcl(i)%>'><%=c_mjcl(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>水接头数量</td>
      <td class=ltd><input type=text name=sjtsl size=30 value=<%=Rs("sjtsl")%>></td>
    </tr>
    <tr>
      <td class=rtd>气接头数量</td>
      <td class=ltd><input type=text name=qjtsl size=30 value=<%=Rs("qjtsl")%>></td>
    </tr>
    <tr>
      <td class=rtd>牵引速度</td>
      <td class=ltd><input type=text name=qysd size=10 value=<%=Rs("qysd")%>>
        米/分(m/min)</td>
    </tr>
    <tr>
      <td class=rtd>挤出方向</td>
      <td class=ltd><select name=jcfx onchange=calmjfz();>
          <option value="/" <%If Rs("jcfx")="/" Then%> selected <%End If%>>&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <option value="技术决定" <%If Rs("jcfx")="技术决定" Then%> selected <%End If%>>技术决定</option>
          <option value="客户决定" <%If Rs("jcfx")="客户决定" Then%> selected <%End If%>>客户决定</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>腔数</td>
      <td class=ltd><input type=text size=4 name=qs onchange=calmjfz(); value=<%=rs("qs")%>>腔</td>
    </tr>
    <tr>
      <td class=rtd>配加热板</td>
      <td class=ltd><select name="pjrb">
          <option value=true<%If Rs("pjrb") Then%> selected<%End If%>>是</option>
          <option value=false<%If not(Rs("pjrb")) Then%> selected<%End If%>>否</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>加热板信息</td>
      <td class=ltd> 相数:
        <select name="jrbxs">
          <option value="两相"<%If Rs("jrbxs")="两相" Then%> selected<%End If%>>两相</option>
          <option value="三相"<%If Rs("jrbxs")="三相" Then%> selected<%End If%>>三相</option>
        </select>
        &nbsp;
        材质:
        <select name="jrbcl">
          <option value="铸铝" <%If Rs("jrbcl")="铸铝" Then%> selected <%End If%>>铸铝</option>
          <option value="云母"<%If Rs("jrbcl")="云母" Then%> selected <%End If%>>云母</option>
        </select>
        &nbsp;
        其它说明:
        <input type=text name=jrbxx size=40 value=<%=Rs("jrbxx")%>></td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>■分值信息■</b></td>
      <td class=ctd></td>
    </tr>
    <tr>
      <td class=rtd>参考断面</td>
      <td class=ltd alt="1.选择参考断面:获得模具的大概分值;<br>2.选择复杂系数:确定模具的具体分值;<br>3.根据不同的模具情况选择:确定模具的最终分值."><select name="ckdm" onChange="if(this.selectedIndex==0) xcsm.innerHTML='';else xcsm.innerHTML=' 分值:'+x_xcfz[this.selectedIndex-1] + '  适合于:' + x_xcsm[this.selectedIndex-1];calmjfz();">
        </select>
        &nbsp;&nbsp; <span id=xcsm></span></td>
    </tr>
    <tr>
      <td class=rtd>定额断面</td>
      <td class=ltd alt="1.选择参考断面:获得模具的基础定额;<br>2.选择复杂系数:确定模具的具体定额;<br>3.根据不同的模具情况选择:确定模具的最终定额."><select name="dedm" onChange="if(this.selectedIndex==0) jcde.innerHTML='';else jcde.innerHTML=' 定额:'+x_defz[this.selectedIndex-1];document.all.defz.value=x_defz[this.selectedIndex-1];">
        </select>
        &nbsp;&nbsp; <span id=jcde></span></td>
    </tr>        
    <tr>
      <td class=rtd>复杂系数</td>
      <td class=ltd><input name=fzxs type=text onchange="calmjfz()" size=5 value=<%=Rs("fzxs")%>></td>
    </tr>
    <tr>
      <td class=rtd>共挤</td>
      <td  class=ltd>
        <input type="checkbox" name="gjfs1" class="radio" value="1" onclick="calmjfz();" <%If NullToNum(Rs("ssgj"))<>0 Then%> checked <%End If%> />
        双色共挤<input name=ssgjf type=text onchange="calmjfz()" value=<%=Rs("ssgj")%> size=5 <%If NullToNum(Rs("ssgj"))=0 Then%> style="display:none" <%End If%>>
        <input type="checkbox" name="gjfs2" class="radio" value="1" onclick="calmjfz();" <%If NullToNum(Rs("qbfgj"))<>0 Then%> checked <%End If%> />
        全包覆共挤<input name=qbfgjf type=text onchange="calmjfz()" value=<%=Rs("qbfgj")%> size=5  <%If NullToNum(Rs("qbfgj"))=0 Then%> style="display:none" <%End If%>>
        <input type="checkbox" name="gjfs3" class="radio" value="1" onclick="document.mtask_add.gjfs4.checked=false;calmjfz();" <%If NullToNum(Rs("qgj"))<>0 Then%> checked <%End If%> />
        软硬前共挤<input name=qgjf type=text onchange="calmjfz()" value=<%=Rs("qgj")%> size=5  <%If NullToNum(Rs("qgj"))=0 Then%> style="display:none" <%End If%>>
        <input type="checkbox" name="gjfs4" class="radio" value="1" onclick="document.mtask_add.gjfs3.checked=false;calmjfz();" <%If NullToNum(Rs("hgj"))<>0 Then%> checked <%End If%> />
        软硬后共挤 <input name=hgjf type=text onchange="calmjfz()" value=<%=Rs("hgj")%> size=5  <%If NullToNum(Rs("hgj"))=0 Then%> style="display:none" <%End If%>> </td>
            </tr>
    <tr>
    <tr>
      <td class=rtd>模具分值</td>
      <td class=ltd> 模具总分:<span id=span_mjzf style="font-weight:bold;"><%=Rs("mjzf")%></span>分&nbsp;&nbsp;&nbsp;&nbsp; <span id=span_gjzf style="font-weight:bold;"> 共挤: <%=Rs("gjzf")%>分</span> <br>
        BOM总分:<span id=span_bomzf style="font-weight:bold;"><%=Rs("bomzf")%></span>分<br>
        调试手册总分:<span id=span_tsdzf style="font-weight:bold;"><%=Rs("tsdzf")%></span>分<br>
        调试总分:<span id=span_tszf style="font-weight:bold;"><%=Rs("tszf")%></span>分<br>
        调试信息整理总分:<span id=span_tsxxzlzf style="font-weight:bold;"><%=Rs("tsxxzlzf")%></span>分<br></td>
    </tr>
    <input type=hidden name=mjzf value=<%=Rs("mjzf")%>>
    <input type=hidden name=gjzf value=<%=Rs("gjzf")%>>
    <input type=hidden name=bomzf value=<%=Rs("bomzf")%>>
    <input type=hidden name=tsdzf value=<%=Rs("tsdzf")%>>
    <input type=hidden name=tszf value=<%=Rs("tszf")%>>
    <input type=hidden name=tsxxzlzf value=<%=Rs("tsxxzlzf")%>>
    <input type=hidden name=bssgj value=false>
    <input type=hidden name=bqbfgj value=false>
    <input type=hidden name=bryqgj value=false>
    <input type=hidden name=bryhgj value=false>
    <input type=hidden name=defz value=0>
    <tr>
      <td class=rtd rowspan="2">分值比例</td>
      <td class=ltd>模头比例:
        <input type=text name=mtbl size=4 onchange=blchange(); value=<%=Rs("mtbl")%>>
        %&nbsp;&nbsp;&nbsp;定型比例:
        <input type=text name=dxbl size=4 disabled value=<%=100-rs("mtbl")%>>
        %</td>
    </tr>
    <tr>
        <td class=ltd>模头结构:
        <input type=text name=mtjgbl size=4 value=<%=Rs("mtjgbl")%>>
        %&nbsp;&nbsp;&nbsp;定型结构:
        <input type=text name=dxjgbl size=4 value=<%=Rs("dxjgbl")%>> %</td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>■模具信息■</b></td>
      <td class=ctd>&nbsp;</td>
    </tr>
    <tr>
      <td class=rtd>模具信息</td>
      <td colspan="2" class=ltd><select name=mjxx onchange='chkmjxx(this);'>
          <option value="全套" <%if rs("mjxx")="全套" then%> selected <%end if%>>全套</option>
          <option value="模头" <%if rs("mjxx")="模头" then%> selected <%end if%>>模头</option>
          <option value="定型" <%if rs("mjxx")="定型" then%> selected <%end if%>>定型</option>
        </select>&nbsp;&nbsp;&nbsp;
        模头<select name=mtrw id="mtrw">
          <option value="" selected></option>
          <option value="设计" <%if rs("mtrw")="设计" then%> selected <%end if%>>设计</option>
          <option value="复改" <%if rs("mtrw")="复改" then%> selected <%end if%>>复改</option>
          <option value="复查" <%if rs("mtrw")="复查" then%> selected <%end if%>>复查</option>
        </select>&nbsp;&nbsp;&nbsp;
        定型<select name=dxrw id="dxrw">
          <option value="" selected></option>
          <option value="设计" <%if rs("dxrw")="设计" then%> selected <%end if%>>设计</option>
          <option value="复改" <%if rs("dxrw")="复改" then%> selected <%end if%>>复改</option>
          <option value="复查" <%if rs("dxrw")="复查" then%> selected <%end if%>>复查</option>
        </select></td>        
    </tr>
    <tr>
      <td class=rtd>任务内容</td>
      <td class=ltd><select name=rwlr>
          <option value="设计" <%If Rs("rwlr")="设计" Then%> selected <%End If%>>设计</option>
          <option value="复改" <%If Rs("rwlr")="复改" Then%> selected <%End If%>>复改</option>
          <option value="复查" <%If Rs("rwlr")="复查" Then%> selected <%End If%>>复查</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>厂内调试</td>
      <td class=ltd><select name="cnts" style="width:51px;" onchange='ExcTslb(this);'>
          <option value=true<%if rs("cnts") then%> selected<%end if%>>是</option>
          <option value=false<%if not(rs("cnts")) then%> selected<%end if%>>否</option>
        </select></td>
    </tr>
    <tr id=trbeit>
      <td class=rtd>北&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;调</td>
      <td class=ltd><select name="beit" style="width:51px;">
          <option value=true <%if rs("beit") then%> selected<%end if%>>是</option>
          <option value=false <%if not(rs("beit")) then%> selected<%end if%>>否</option>
        </select></td>
    </tr>
    <tr id=trtslb>
      <td class=rtd>调试类别</td>
      <td class=ltd alt="1.选择型材类别:确定模具的最多可调试次数；<br>2.实际次数与最大次数的差说明模具结构及调试方案的正确性."><select name="tslb" style="width:51px;" onChange="if(this.selectedIndex==0) xcts.innerHTML='';else xcts.innerHTML=' 额定调试次数:'+z_xccs[this.selectedIndex-1] + ' - ' + z_xcfw[this.selectedIndex-1] + '  适用于:' + z_xcbz[this.selectedIndex-1];">
        </select>
        &nbsp;&nbsp; <span id=xcts></span></td>
    </tr>
    <tr>
      <td class=rtd>型材壁厚</td>
      <td class=ltd><input type=text name=xcbh size=14 value=<%=Rs("xcbh")%>></td>
    </tr>
    <tr id=trdxqg>
      <td class=rtd>定型切割</td>
      <td class=ltd><select name="dxqg">
          <option value="不合割"<%If Rs("dxqg")="不合割" Then%> selected<%End If%>>不合割</option>
          <option value="分体合割"<%If Rs("dxqg")="分体合割" Then%> selected<%End If%>>分体合割</option>
          <option value="整体合割"<%If Rs("dxqg")="整体合割" Then%> selected<%End If%>>整体合割</option>
          <option value="普线一次切割"<%if rs("dxqg")="普线一次切割" then%> selected<%end if%>>普线一次切割
        </select></td>
    </tr>
    <tr id=trmtjg>
      <td class=rtd>模头结构</td>
      <td class=ltd><input type=text name=mtjg size=30 value=<%=Rs("mtjg")%>>
        </td>
    </tr>
    <tr id=trdxjg>
      <td class=rtd>定型结构</td>
      <td class=ltd><input type=text name=dxjg size=30 value=<%=Rs("dxjg")%>>
        &nbsp;
        <select onchange='this.form.dxjg.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_dxjg)%>
          <option value='<%=c_dxjg(i)%>'><%=c_dxjg(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr id=trsxjg>
      <td class=rtd>水箱结构</td>
      <td class=ltd><input type=text name=sxjg size=30 value=<%=Rs("sxjg")%>>
        &nbsp;
        <select onchange='this.form.sxjg.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_sxjg)%>
          <option value='<%=c_sxjg(i)%>'><%=c_sxjg(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>模头连接尺寸</td>
      <td class=ltd><input type=text name=mtljcc size=40 value="<%=Rs("mtljcc")%>">
        &nbsp;
        <select onchange='this.form.mtljcc.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_mtljcc)%>
          <option value='<%=c_mtljcc(i)%>'><%=c_mtljcc(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>热电偶规格</td>
      <td class=ltd><input type=text name=rdogg size=30 value="<%=Rs("rdogg")%>">
        &nbsp;
        <select onchange='this.form.rdogg.value=this.value;'>
          <option></option>
          "
          <%for i = 0 to ubound(c_rdogg)%>
          <option value='<%=c_rdogg(i)%>'><%=c_rdogg(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>■其他信息■</b></td>
      <td class=ctd>&nbsp;</td>
    </tr>
    <tr>
      <td class=rtd>备注</td>
      <td class=ltd><textarea name="bz" cols="75" rows="7"><%=Rs("bz")%></textarea></td>
    </tr>
    <tr>
      <td class=rtd>计划开始时间</td>
      <td class=ltd><%
		dim Tmpshsj
		If IsNull(Rs("jhkssj")) Then Tmpshsj=now else Tmpshsj=Rs("jhkssj")
			%>
        <script language=javascript>
 		var myDate=new dateSelector(<%=year(Tmpshsj)&","&month(Tmpshsj)&","&day(Tmpshsj)%>);
  		myDate.year;
 		myDate.inputName='jhkssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>计划结构结束时间</td>
      <td class=ltd><%
		dim Tmpjssj
		If IsNull(Rs("jhjgsj")) Then Tmpjssj=Rs("jhjssj") else Tmpjssj=Rs("jhjgsj")
			%>
        <script language=javascript>
 		var myDate=new dateSelector(<%=year(Tmpjssj)&","&month(Tmpjssj)&","&day(Tmpjssj)%>);
  		myDate.year;
 		myDate.inputName='jgjssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>计划全套结束时间</td>
      <td class=ltd><%
		If IsNull(Rs("jhjssj")) Then Tmpjssj=now else Tmpjssj=Rs("jhjssj")
			%>
        <script language=javascript>
 		var myDate=new dateSelector(<%=year(Tmpjssj)&","&month(Tmpjssj)&","&day(Tmpjssj)%>);
  		myDate.year;
 		myDate.inputName='jhjssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script></td>
    </tr>
    <% '13:40 2007-1-6-星期六
'兼容以前只有一个任务只有一个组长的版本
If Rs("zz") <> "" Then %>
    <tr>
      <td class=rtd>组长</td>
      <td class=ltd><select name="zz">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%If Rs("zz")=c_allzz(i) Then%> selected<%End If%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <%else%>
    <tr>
      <td class=rtd>结构组长</td>
      <td class=ltd><select name="jgzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%if rs("jgzz")=c_allzz(i) then%> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>设计组长</td>
      <td class=ltd><select name="sjzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%if rs("sjzz")=c_allzz(i) then%> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <%End If%>
    <tr>
      <td class=rtd>技术代表</td>
      <td class=ltd><select name="jsdb" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'<%If Rs("jsdb")=c_allzy(i) Then%> selected<%End If%>><%=c_allzy(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=ctd colspan=2><input type=submit value=" · 确 定 · "></td>
    </tr>
    <input type="hidden" name=id value=<%=Rs("id")%>>
  </form>
</table>
<%
	Dim TmpTslb,TmpCkdm,TmpDmde
	TmpTslb=Rs("tslb")
	TmpCkdm=Rs("ckdm")
	TmpDmde=Rs("dedm")
	Call mtask_js(TmpCkdm,TmpTslb,TmpDmde)
End Function		'mtask_change()

Function BzChan(Rs)
%>
<%Call TbTopic("更改流水号 <font style=color:#0000FF>" & Rs("lsh") & "</font> 的备注")%>
<%If ChkAble(4) Then Response.Write "<a href=mtask_zzchange.asp?s_lsh="&Rs("lsh")&">组长权限</a><br>"%>
<table class=xtable cellspacing=0 cellpadding=10>
  <form id=Bz_change name=Bz_change action=mtask_indb.asp?action=BzChan method=post onSubmit='return true;'>
	<tr>
		<td class=ctd >订单号</td>
		<td class=ctd ><%=Rs("ddh")%></td>
		<td class=ctd >客户名</td>
		<td class=ctd ><%=Rs("dwmc")%></td>
	</tr>
    <tr>
      <td class=ctd colspan="4"><textarea name="bz" cols="75" rows="10"><%=Rs("bz")%></textarea></td>
    </tr>
    <tr>
      <td class=ctd colspan="4"><input type=submit value=" · 确 定 · "></td>
    </tr>
    <input type="hidden" name=id value=<%=Rs("id")%>>
  </form>
</table>
<%
End Function		'BzChan()

Function mtask_js(CkdmOv,TslbOv,DmdeOV)
'以下为JS代码%>
<script language="javascript">
//初始化北调与调试类别的出现与否
<%If Rs("cnts") Then%>
	document.all.trtslb.style.display="";
	document.all.trbeit.style.display="none";
<%else%>
	document.all.trtslb.style.display="none";
	document.all.trbeit.style.display="";
<%End If%>

//厂内调试下拉菜单关联动作
function ExcTslb(ChangeV)
{
	if (ChangeV.value=="true")
		{
			document.all.trtslb.style.display="";
			document.all.trbeit.style.display="none";
		}
	else
		{
			document.all.trtslb.style.display="none";
			document.all.trbeit.style.display="";
		}
}

//对参考模具初始化
	var x_xcmc = new Array();
	var x_xcfz = new Array();
	var x_xcsm = new Array();

<%
	set Rs=xjweb.exec("select * from c_dmfz order by dmmc",1)
	i=0
	do while not Rs.Eof
%>
		x_xcmc[<%=i%>]="<%=Rs("dmmc")%>";
		x_xcfz[<%=i%>]="<%=Rs("dmfz")%>";
		x_xcsm[<%=i%>]="<%=Rs("bz")%>";
<%
		i = i + 1
		Rs.movenext
	loop
%>
	for(var i=1; i<x_xcmc.length + 1; i++)
	{
		document.all.ckdm[i] = new Option(x_xcmc[i-1],x_xcmc[i-1]);
		if(document.all.ckdm.options[i].value=="<%=CkdmOV%>")
 			document.all.ckdm.options[i].selected=true; 
	}
	calmjfz();
//对定额断面初始化
	var x_demc = new Array();
	var x_defz = new Array();

<%
	set rs=xjweb.exec("select * from c_dmde",1)
	i=0
	do while not rs.eof
%>
		x_demc[<%=i%>]="<%=rs("dmmc")%>";
		x_defz[<%=i%>]="<%=rs("dmfz")%>";
<%
		i = i + 1
		rs.movenext
	loop
	rs.close
%>
	for(var i=1; i<x_demc.length + 1; i++)
	{
		document.all.dedm[i] = new Option(x_demc[i-1],x_demc[i-1]);
		if(document.all.dedm.options[i].value=="<%=DmdeOV%>")
		{
 			document.all.dedm.options[i].selected=true; 
			document.all.defz.value=x_defz[i-1];
		}
	}

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
	}
 	var   o   =   document.getElementById("tslb");
	var   i;
	for(i=0;i<o.length;i++)
	{
 		if(o.options[i].value=="<%=TslbOv%>")
 			o.options[i].selected=true;
 	}

//计算模具分值,同时显示隐藏层
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

	var ttmjfz=0;		//模具总分值
	var ttgjfz=0;		//共挤分值
	var tmpobj;
	var tmpstr;
	document.all.span_gjzf.innerHTML="";
	if(isNaN(parseFloat(document.all.mtjgbl.value))) document.all.mtjgbl.value=mtjgbl*100;
	if(isNaN(parseFloat(document.all.dxjgbl.value))) document.all.dxjgbl.value=dxjgbl*100;

	//各项参数的值
	var i_gjfs=0;
	var i_qhgj=0;

	var str=document.all;
	//由参考断面获得初始分值
	if((str.ckdm.selectedIndex-1)>=0) ttmjfz=x_xcfz[str.ckdm.selectedIndex-1];

	//共挤确定;
	var issgjf=str.ssgjf.value*1;
	var iqbfgjf=str.qbfgjf.value*1;
	var iqgjf=str.qgjf.value*1;
	var ihgjf=str.hgjf.value*1;
	if (str.gjfs1.checked)	//双色共挤
	{
		if (issgjf>0)
		{
			ttgjfz=issgjf;
		}
		else
		{
			ttgjfz=ssgjfz;
			str.ssgjf.value=ssgjfz;
		}
		str.ssgjf.style.display='';
		str.span_gjzf.style.display='';
		str.span_gjzf.innerHTML="共挤: " + Math.round(ttgjfz) + "分"
	}
	else
	{
		str.ssgjf.value=0;
		str.ssgjf.style.display="none";
	}
	if (str.gjfs2.checked)	//全包覆共挤
	{
		if (iqbfgjf>0)
		{
			ttgjfz=ttgjfz+iqbfgjf;
		}
		else
		{
			ttgjfz=ttgjfz+qbfgjfz;
			str.qbfgjf.value=qbfgjfz;
		}
		str.qbfgjf.style.display='';
		str.span_gjzf.style.display='';
		str.span_gjzf.innerHTML="共挤: " + Math.round(ttgjfz) + "分"
	}
	else
	{
		str.qbfgjf.value=0;
		str.qbfgjf.style.display="none";
	}
	if (str.gjfs3.checked)	//软硬前共挤
	{
		if (iqgjf>0)
		{
			ttgjfz=ttgjfz+iqgjf;
		}
		else
		{
			ttgjfz=ttgjfz+ryqgjfz;
			str.qgjf.value=ryqgjfz;
		}
		str.qgjf.style.display='';
		str.span_gjzf.style.display='';
		str.span_gjzf.innerHTML="共挤: " + Math.round(ttgjfz) + "分"
	}
	else
	{
		str.qgjf.value=0;
		str.qgjf.style.display="none";
	}
	if (str.gjfs4.checked)	//软硬后共挤
	{
		if (ihgjf>0)
		{
			ttgjfz=ttgjfz+ihgjf;
		}
		else
		{
			ttgjfz=ttgjfz+ryhgjfz;
			str.hgjf.value=ryhgjfz;
		}
		str.hgjf.style.display='';
		str.span_gjzf.style.display='';
		str.span_gjzf.innerHTML="共挤: " + Math.round(ttgjfz) + "分"
	}
	else
	{
		str.hgjf.value=0;
		str.hgjf.style.display="none";
	}

	//腔数确定
	//ttmjfz=ttmjfz*(Math.sqrt(str.qs.value));

	//复杂系数确定
	ttmjfz=ttmjfz*(str.fzxs.value);

	//模具总分
	var ttmjfz=ttmjfz+ttgjfz;

	//str=document.all
	str.span_mjzf.innerHTML=Math.round(ttmjfz);
	var tmpmjfz=0;

	tmpmjfz=ttmjfz;
	str.span_bomzf.innerHTML=Math.round(tmpmjfz*bomxs);
	str.span_tsdzf.innerHTML=Math.round(tmpmjfz*tsdxs);
	str.span_tszf.innerHTML=Math.round(tmpmjfz*tsxs);
	str.span_tsxxzlzf.innerHTML=Math.round(tmpmjfz*tsxxzlxs);

	str.mjzf.value=Math.round(ttmjfz);
	str.gjzf.value=Math.round(ttgjfz);
	str.bomzf.value=Math.round(tmpmjfz*bomxs);
	str.tsdzf.value=Math.round(tmpmjfz*tsdxs);
	str.tszf.value=Math.round(tmpmjfz*tsxs);
	str.tsxxzlzf.value=Math.round(tmpmjfz*tsxxzlxs);
	//document.all.mjzf.innerHTML=Math.round(ttmjfz);
}

//模头分值比例
function blchange()
{
	//f!(isNaN(document.all.mtbl.value)) document.all.dxbl.value=30;
	if(document.all.mtbl.disabled==true) return;
	var mtbl=0;
	var dxbl=0;

	if(!isNaN(parseFloat(document.all.mtbl.value))) mtbl=parseFloat(document.all.mtbl.value);
	if(!isNaN(parseFloat(document.all.dxbl.value))) dxbl=parseFloat(document.all.dxbl.value);
	if(mtbl>100) mtbl=100;
	if(dxbl>100) dxbl=100;
	if(mtbl<0) mtbl=0;
	if(dxbl<0) dxbl=0

	dxbl=100-mtbl;
	document.all.dxbl.value=dxbl;
	document.all.mtbl.value=mtbl;
	return;
}

//模具信息
function chkmjxx(ftemp)
{
	switch (ftemp.value)
	{
		case "模头" :
			document.all.trdxqg.style.display="none";
			document.all.trdxjg.style.display="none";
			document.all.trsxjg.style.display="none";
			document.all.dxqg.value="";
			document.all.dxjg.value="/";
			document.all.sxjg.value="/";
			break;
		default:
			document.all.trdxqg.style.display="";
			document.all.trdxjg.style.display="";
			document.all.trsxjg.style.display="";
			break;
	}
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
	//alert(igfdm[1]);
	document.all.gfdm.length = 0;
	document.all.gfdm.options[document.all.gfdm.length] = new Option("请选择","");
	for (j=0;j<igfdm.length;j++){
		if(igfdm[j][0]==selvalue){
			document.all.gfdm.options[document.all.gfdm.length] = new Option(igfdm[j][1],igfdm[j][1]);
		}
	}
}
document.all.gfdm.options[document.all.gfdm.length] = new Option("请选择","");
calmjfz();
chkmjxx(document.all.mjxx);
</script>
<%
End Function	'mtask_js()
%>
