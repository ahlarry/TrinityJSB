<!--#include file="include/conn.asp"-->
<%
Call ChkAdminAble()
xjweb.Header()
dim action
action=request("action")
select case action
	case "depart"
		call common_dis("depart", "用户部门")
	case "depart_addindb"
		call common_add("depart", "用户部门")
	case "depart_delindb"
		call common_del("depart", "用户部门")

	case "sbcj"
		call common_dis("sbcj", "设备厂家")
	case "sbcj_addindb"
		call common_add("sbcj", "设备厂家")
	case "sbcj_delindb"
		call common_del("sbcj", "设备厂家")

	case "dwmc"
		call common_dis("dwmc", "单位名称")
	case "dwmc_addindb"
		call common_add("dwmc", "单位名称")
	case "dwmc_delindb"
		call common_del("dwmc", "单位名称")
	
	case "mtjg"
		call common_dis("mtjg", "模头结构")
	case "mtjg_addindb"
		call common_add("mtjg", "模头结构")
	case "mtjg_delindb"
		call common_del("mtjg", "模头结构")
	
	case "dxjg"
		call common_dis("dxjg", "定型结构")
	case "dxjg_addindb"
		call common_add("dxjg", "定型结构")
	case "dxjg_delindb"
		call common_del("dxjg", "定型结构")

	case "sxjg"
		call common_dis("sxjg", "水箱结构")
	case "sxjg_addindb"
		call common_add("sxjg", "水箱结构")
	case "sxjg_delindb"
		call common_del("sxjg", "水箱结构")

	case "dmmc"
		call common_dis("dmmc", "断面名称")
	case "dmmc_addindb"
		call common_add("dmmc", "断面名称")
	case "dmmc_delindb"
		call common_del("dmmc", "断面名称")

	case "mjcl"
		call common_dis("mjcl", "模具材料")
	case "mjcl_addindb"
		call common_add("mjcl", "模具材料")
	case "mjcl_delindb"
		call common_del("mjcl", "模具材料")

	case "jcjxh"
		call common_dis("jcjxh", "挤出机型号")
	case "jcjxh_addindb"
		call common_add("jcjxh", "挤出机型号")
	case "jcjxh_delindb"
		call common_del("jcjxh", "挤出机型号")

	case "rdogg"
		call common_dis("rdogg", "热电偶规格")
	case "rdogg_addindb"
		call common_add("rdogg", "热电偶规格")
	case "rdogg_delindb"
		call common_del("rdogg", "热电偶规格")

	case "mtljcc"
		call common_dis("mtljcc", "模头连接尺寸")
	case "mtljcc_addindb"
		call common_add("mtljcc", "模头连接尺寸")
	case "mtljcc_delindb"
		call common_del("mtljcc", "模头连接尺寸")

	case "lxrwlx"
		call common_dis("lxrwlx", "零星任务类型")
	case "lxrwlx_addindb"
		call common_add("lxrwlx", "零星任务类型")
	case "lxrwlx_delindb"
		call common_del("lxrwlx", "零星任务类型")

	case "fzbl"
		call fzbl_dis()
	case "fzbl_indb"
		call fzbl_indb()

	case "ckfz"
		call ckfz_dis()
	case "ckfz_addindb"
		call ckfz_addindb()
	case "ckfz_chgindb"
		call ckfz_chgindb()
	case "ckfz_delindb"
		call ckfz_delindb()

	case else
		call main()
end select
Call xjweb.Footer()

function main()
	Call TbTopoic("请从左边选择您要操作的项目!")
end function

'-------------------------------------------------------公用函数------------------------------------------------------------------------------
rem 公用函数
function common_dis(item, stritem)		'显示界面   item 为代号, stritem-- 为名称   如item=dwmc, stritem="单位名称"
	Call TbTopic(web_info(0) & "<font style=color:#ff0000;>"& stritem &"</font> -- 查询管理")
%>
	<table border=0 cellspacing=0 cellpadding=3 class=xtable width="60%">
		<form action="<%=request.servervariables("script_name")%>?action=<%=item%>_addindb" method="post" onsubmit="return chkaddinf();">
		<tr>
			<td class=rtd>添加<%=stritem%></td>
			<td class=ltd><input type=text name="add<%=item%>"></td>
			<td class=ctd><input type=submit value=" 添加 "></td>
		</tr>
		
		</form>
		<form action="<%=request.servervariables("script_name")%>?action=<%=item%>_delindb" method="post" onsubmit="return chkdelinf();">
		<tr>
			<td class=rtd>删除<%=stritem%></td>
			<td class=ltd>
				<select name="del<%=item%>"><option value=""></option>
			<%
				strSql="select * from [c_"&item&"] order by "&item&""
				set rs=xjweb.Exec(strSql,1)
				do while not rs.eof
			%>
					<option value="<%=rs(item)%>"><%=rs(item)%></option>
			<%
					rs.movenext
				loop
				rs.close
			%>
			</td>
			<td class=ctd><input type=submit value=" 删除 "></td>
		</tr>
		</form>
	</table>
	<script language="javascript">
	function chkaddinf()
	{
		if(document.all.add<%=item%>.value==""){alert('请输入要添加的<%=stritem%>!');document.all.add<%=item%>.focus();return false;}
	}
	function chkdelinf()
	{
		if(document.all.del<%=item%>.value=="")
			{alert('请选择要删除的<%=stritem%>!');document.all.del<%=item%>.focus();return false;}
		else
			{return confirm('确信要删除<%=stritem%> '+document.all.del<%=item%>.value+' 吗?');}
	}
	</script>
<%
end function

'---------------------------------------------添加入库函数-------------------------------------------------------------
function common_add(item, stritem)
	dim addinfo
	addinfo=trim(request("add"&item))
	if addinfo="" then response.write(prompt(stritem & "不能为空!")) : exit function
	strSql="select "&item&" from [c_"&item&"] where "&item&"='"&addinfo&"'"
	set rs=xjweb.Exec(strSql, 1)
	if not(rs.eof or rs.bof) then Call JsAlert(""&stritem & " " & addinfo & " 已经存在!","?action="&item&"") : Exit Function
	rs.close
	strSql="insert into c_"&item&" ("&item&") values ('"&addinfo&"')"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " 添加成功!"))
	Call JsAlert(""&stritem & " " & addinfo & " 添加成功!","?action="&item&"")
end function
'--------------------------------------------------从库中删除函数-----------------------------------------------------------
Function common_del(item, stritem)
	dim delinfo
	delinfo=trim(request("del"&item))
	if delinfo="" then response.write(prompt("请选择要删除的"&stritem&"!")) : exit function
	strSql="delete from c_"&item&" where "&item&"='"&delinfo&"'"
	call xjweb.Exec(strSql, 0)
	'response.write(prompt(stritem & " " & delinfo & " 删除成功!"))
	Call JsAlert(""&stritem & " " & delinfo & " 删除成功!","?action="&item&"")
End Function
'----------------------------------------------------公用结束-----------------------------------------------------


'------------------------------------------------------分值比例开始--------------------------------------------------
function fzbl_dis()		'分值比例
	dim fzbl(15), fzblsm(15)
	strSql="select * from c_fzbl"
	set rs=xjweb.Exec(strSql,1)
	if rs.eof or rs.bof then rs.close : response.write ("系统故障") : exit function
		fzbl(0)=rs("jgbl")*100
		fzbl(1)=rs("sjbl")*100
		fzbl(2)=rs("shbl")*100
		fzbl(3)=rs("fgbl")*100
		fzbl(4)=rs("fgshbl")*100
		fzbl(5)=rs("fcbl")*100
		fzbl(6)=rs("gwfzxs")*100
		fzbl(7)=rs("msfzxs")*100
		fzbl(8)=rs("hgjfzxs")*100
		fzbl(9)=rs("ssgjfzxs")*100
		fzbl(10)=rs("rygjfzxs")*100
		fzbl(11)=rs("bomfzxs")*100
		fzbl(12)=rs("tsdfzxs")*100
		fzbl(13)=rs("tsfzxs")*100
		fzbl(14)=rs("tsxxzlfzxs")*100
	rs.close
	fzblsm(0)="模具分值"
	fzblsm(1)="模具分值"
	fzblsm(2)="模具分值"
	fzblsm(3)="模具分值"
	fzblsm(4)="模具分值"
	fzblsm(5)="模具分值"
	fzblsm(6)="模具分值"
	fzblsm(7)="模具分值"
	fzblsm(8)="模具分值"
	fzblsm(9)="模具分值"
	fzblsm(10)="模具分值"
	fzblsm(11)="模具分值"
	fzblsm(12)="模具分值"
	fzblsm(13)="模具分值"
	fzblsm(14)="模具分值"
	Call TbTopic(web_info(0) & " <font style=color:#ff0000;>分值比例(系数)</font> -- 查询管理")
%>
	<table border=0 cellspacing=0 cellpadding=3 class=xtable width="60%">
		<form name="fzbl" action="<%=request.servervariables("script_name")%>?action=fzbl_indb" method="post" onsubmit="return chkinf();">
		<tr><td class=ctd colspan=3><b>分值比例(系数)</b></td></tr>
		<tr>
			<td class=rtd width="40%">结构比例</td>
			<td class=ltd width="*"><input type=text name=jgbl size=6 value=<%=fzbl(0)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(0)%>"></td>
		</tr>
		<tr>
			<td class=rtd>设计比例</td>
			<td class=ltd><input type=text name=sjbl size=6 value=<%=fzbl(1)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(1)%>"></td>
		</tr>
		<tr>
			<td class=rtd>审核比例</td>
			<td class=ltd><input type=text name=shbl size=6 value=<%=fzbl(2)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(2)%>"></td>
		</tr>
		<tr>
			<td class=rtd>复改比例</td>
			<td class=ltd><input type=text name=fgbl size=6 value=<%=fzbl(3)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(3)%>"></td>
		</tr>
		<tr>
			<td class=rtd>复改审核比例</td>
			<td class=ltd><input type=text name=fgshbl size=6 value=<%=fzbl(4)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(4)%>"></td>
		</tr>
		<tr>
			<td class=rtd>复查比例</td>
			<td class=ltd><input type=text name=fcbl size=6 value=<%=fzbl(5)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(5)%>"></td>
		</tr>
		<tr>
			<td class=rtd>国外分值系数</td>
			<td class=ltd><input type=text name=gwfzxs size=6 value=<%=fzbl(6)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(6)%>"></td>
		</tr>
		<tr>
			<td class=rtd>美式分值系数</td>
			<td class=ltd><input type=text name=msfzxs  size=6 value=<%=fzbl(7)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(7)%>"></td>
		</tr>
		<tr>
			<td class=rtd>后共挤分值系数</td>
			<td class=ltd><input type=text name=hgjfzxs size=6 value=<%=fzbl(8)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(8)%>"></td>
		</tr>
		<tr>
			<td class=rtd>双色共挤分值系数</td>
			<td class=ltd><input type=text name=ssgjfzxs size=6 value=<%=fzbl(9)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(9)%>"></td>
		</tr>
		<tr>
			<td class=rtd>软硬共挤分值系数</td>
			<td class=ltd><input type=text name=rygjfzxs size=6 value=<%=fzbl(10)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(10)%>"></td>
		</tr>
		<tr>
			<td class=rtd>BOM分值系数</td>
			<td class=ltd><input type=text name=bomfzxs  size=6 value=<%=fzbl(11)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(11)%>"></td>
		</tr>
		<tr>
			<td class=rtd>调试单分值系数</td>
			<td class=ltd><input type=text name=tsdfzxs size=6 value=<%=fzbl(12)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(12)%>"></td>
		</tr>
		<tr>
			<td class=rtd>调试分值系数</td>
			<td class=ltd><input type=text name=tsfzxs size=6 value=<%=fzbl(13)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(13)%>"></td>
		</tr>
		<tr>
			<td class=rtd>调试信息整理分值系数</td>
			<td class=ltd><input type=text name=tsxxzlfzxs size=6 value=<%=fzbl(14)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(14)%>"></td>
		</tr>
		<tr><td class=ctd colspan=3><input type=submit value=" 更改 "></td></tr>
		</form>
	</table>
	<script language="javascript">
	function chkinf()
	{
		var frm=document.all.fzbl;
		var jgbl, sjbl, shbl, fgbl, fgshbl;
		var fcbl, gwfzxs, msfzxs, hgjfzxs, ssgjfzxs;
		var rygjfzxs, bomfzxs, tsdfzxs, tsfzxs, tsxxzlfzxs;
		<%
			call jschk("jgbl", "结构比例", fzbl(0))
			call jschk("sjbl", "设计比例", fzbl(1))
			call jschk("shbl", "审核比例", fzbl(2))
			call jschk("fgbl", "复改比例", fzbl(3))
			call jschk("fgshbl", "复改审核比例", fzbl(4))
			call jschk("fcbl", "复查比例", fzbl(5))
			call jschk("gwfzxs", "国外分值系数", fzbl(6))
			call jschk("msfzxs", "美式分值系数", fzbl(7))
			call jschk("hgjfzxs", "后共挤分值系数", fzbl(8))
			call jschk("ssgjfzxs", "双色共挤分值系数", fzbl(9))
			call jschk("rygjfzxs", "软硬共挤分值系数", fzbl(10))
			call jschk("bomfzxs", "BOM分值系数", fzbl(11))
			call jschk("tsdfzxs", "调试单分值系数", fzbl(12))
			call jschk("tsfzxs", "调试分值系数", fzbl(13))
			call jschk("tsxxzlfzxs", "调试信息整理分值系数", fzbl(14))			
		%>
	}
	</script>
<%
end function

function jschk(item, stritem, i)
%>
	if(isNaN(parseFloat(frm.<%=item%>.value))) 
		{alert('<%=stritem%>必须为数字!');frm.<%=item%>.value=<%=i%>;frm.<%=item%>.focus();return false;}
	frm.<%=item%>.value=parseFloat(frm.<%=item%>.value);
	if(<%=item%>>100 || <%=item%><0)	{alert('<%=stritem%>必须在 0 和 100 之间!');frm.<%=item%>.value=<%=i%>;frm.<%=item%>.focus();return false;}
<%
end function

'---------------------------------------------添加入库函数-------------------------------------------------------------
function fzbl_indb()
	dim fzbl(15)
	fzbl(0)=csng(request("jgbl"))/100
	fzbl(1)=csng(request("sjbl"))/100
	fzbl(2)=csng(request("shbl"))/100
	fzbl(3)=csng(request("fgbl"))/100
	fzbl(4)=csng(request("fgshbl"))/100
	fzbl(5)=csng(request("fcbl"))/100
	fzbl(6)=csng(request("gwfzxs"))/100
	fzbl(7)=csng(request("msfzxs"))/100
	fzbl(8)=csng(request("hgjfzxs"))/100
	fzbl(9)=csng(request("ssgjfzxs"))/100
	fzbl(10)=csng(request("rygjfzxs"))/100
	fzbl(11)=csng(request("bomfzxs"))/100
	fzbl(12)=csng(request("tsdfzxs"))/100
	fzbl(13)=csng(request("tsfzxs"))/100
	fzbl(14)=csng(request("tsxxzlfzxs"))/100

	strSql="select * from c_fzbl"
	Call xjweb.Exec("",-1)
	'set rs=server.createobject("adodb.recordset")
	rs.open strSql, conn, 1, 3
		rs("jgbl")=fzbl(0)
		rs("sjbl")=fzbl(1)
		rs("shbl")=fzbl(2)
		rs("fgbl")=fzbl(3)
		rs("fgshbl")=fzbl(4)
		rs("fcbl")=fzbl(5)
		rs("gwfzxs")=fzbl(6)
		rs("msfzxs")=fzbl(7)
		rs("hgjfzxs")=fzbl(8)
		rs("ssgjfzxs")=fzbl(9)
		rs("rygjfzxs")=fzbl(10)
		rs("bomfzxs")=fzbl(11)
		rs("tsdfzxs")=fzbl(12)
		rs("tsfzxs")=fzbl(13)
		rs("tsxxzlfzxs")=fzbl(14)
	rs.update
	rs.close
	set rs=nothing
	Call JsAlert("分值比例(系数) 更改成功!","?action=fzbl")
end function
'----------------------------------------------------分值比例结束-----------------------------------------------------


'------------------------------------------------------参考分值开始--------------------------------------------------
function ckfz_dis()		'参考分值
	Call TbTopic(web_info(0) & " <font style=color:#ff0000;>参考分值</font> -- 查询管理")
%>
	<table border=0 cellspacing=0 cellpadding=3 class=xtable width="80%">
		<form action="<%=request.servervariables("script_name")%>?action=ckfz_addindb" method="post" onsubmit="return chkaddinf();">
		<tr>
			<td class=rtd rowspan=3>添加参考断面</td>
			<td class=rtd>断面名称:</td>
			<td class=ltd><input type=text name="addckdm"></td>
			<td class=ctd rowspan=3><input type=submit value=" 添加 "></td>
		</tr>
		<tr>
			<td class=rtd>断面分值:</td>
			<td class=ltd><input type=text size=10 name="adddmfz"></td>
		</tr>
		<tr>
			<td class=rtd>适用断面:</td>
			<td class=ltd><input type="text" size="60" name="addbz"></td></td>
		</tr>
		</form>
		<form action="<%=request.servervariables("script_name")%>?action=ckfz_chgindb" method="post" onsubmit="return chkchginf();">
		<tr>
			<td class=rtd rowspan=3>更改参考断面</td>
			<td class=rtd>断面名称:</td>
			<td class=ltd>
				<select name="chgckdm" onchange="if(this.selectedIndex==0){chgdmfz.value='';chgbz.value='';}else {chgdmfz.value=x_xcfz[this.selectedIndex-1]; chgbz.value=x_xcsm[this.selectedIndex-1];}">
				</select>
			</td>
			<td class=ctd rowspan=3><input type=submit value=" 更改 "></td>
		</tr>
		<tr>
			<td class=rtd>断面分值:</td>
			<td class=ltd><input type=text size=10 name="chgdmfz"></td>
		</tr>
		<tr>
			<td class=rtd>适用断面:</td>
			<td class=ltd><input type="text" size="60" name="chgbz"></td>
		</tr>
		</form>
		<form action="<%=request.servervariables("script_name")%>?action=ckfz_delindb" method="post" onsubmit="return chkdelinf();">
		<tr>
			<td class=rtd rowspan=3>删除参考断面</td>
			<td class=rtd>断面名称:</td>
			<td class=ltd>
				<select name="delckdm" onchange="if(this.selectedIndex==0){deldmfz.innerHTML='';delbz.innerHTML='';}else {deldmfz.innerHTML=x_xcfz[this.selectedIndex-1]; delbz.innerHTML=x_xcsm[this.selectedIndex-1];}">
				</select>
			</td>
			<td class=ctd rowspan=3><input type=submit value=" 删除 "></td>
		</tr>
		<tr>
			<td class=rtd>断面分值:</td>
			<td class=ltd><span id="deldmfz"></span>&nbsp;</td>
		</tr>
		<tr>
			<td class=rtd>适用断面:</td>
			<td class=ltd><span id="delbz"></span>&nbsp;</td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		//对参考模具初始化
		var x_xcmc = new Array();
		var x_xcfz = new Array();
		var x_xcsm = new Array();
<%
		strSql="select * from c_dmfz order by dmmc"
		set rs=xjweb.Exec(strSql, 1)
		i=0
		do while not rs.eof
%>
			x_xcmc[<%=i%>]="<%=rs("dmmc")%>";
			x_xcfz[<%=i%>]="<%=rs("dmfz")%>";
			x_xcsm[<%=i%>]="<%=rs("bz")%>";
			//response.Write("document.all.XCMC[0] = new Option("&rs("bz")&","&rs("bz")&");")
<%
			i = i + 1
			rs.movenext
		loop
%>
		for(var i=1; i<x_xcmc.length + 1; i++)
	{
		document.all.chgckdm[i] = new Option(x_xcmc[i-1],x_xcmc[i-1]);
		document.all.delckdm[i] = new Option(x_xcmc[i-1],x_xcmc[i-1]);
	}
	//document.all.chgckdm[1] = new Option("xujian","xujian");
	function chkaddinf()
	{
		if(document.all.addckdm.value=="")	{alert('断面名称不能为空!');document.all.addckdm.focus();return false;}	if(isNaN(parseFloat(document.all.adddmfz.value))){alert('断面分值只能为数字!');document.all.adddmfz.value='';document.all.adddmfz.focus();return false;}
		document.all.adddmfz.value=parseFloat(document.all.adddmfz.value)
		if(document.all.addbz.value=="")	{alert('适用断面不能为空!');document.all.addbz.focus();return false;}
	}

	function chkchginf()
	{
		if(document.all.chgckdm.value=="")	{alert('请选择要更改断面的名称!');document.all.chgckdm.focus();return false;}
		if(isNaN(parseFloat(document.all.chgdmfz.value))){alert('断面分值只能为数字!');document.all.chgdmfz.value='';document.all.chgdmfz.focus();return false;}
		document.all.chgdmfz.value=parseFloat(document.all.chgdmfz.value)
		if(document.all.chgbz.value=="")	{alert('适用断面不能为空!');document.all.chgbz.focus();return false;}
	}
	function chkdelinf()
	{
		if(document.all.delckdm.value=="")	
			{alert('请选择要删除断面的名称!');document.all.delckdm.focus();return false;}
		else
			{return confirm('确信要删除参考断面 '+document.all.delckdm.value+' 吗?')}
	}
	</script>
<%
end function

function ckfz_addindb()
	dim dmmc, dmfz, bz
	dmmc="" : dmfz=0 : bz=""
	dmmc=trim(request("addckdm"))
	dmfz=trim(request("adddmfz"))
	bz=trim(request("addbz"))
	if dmmc="" or dmfz="" or dmfz=0 or bz="" then response.write("数据不完整!请从正确入口进入!") : exit function
	strSql="select dmmc from c_dmfz where dmmc='"&dmmc&"'"
	set rs=xjweb.Exec(strSql, 1)
	if not(rs.eof or rs.bof) then Call JsAlert("参考断面 "&dmmc & " 已经存在!","?action=ckfz") : Exit Function
	rs.close
	strSql="insert into c_dmfz (dmmc, dmfz, bz) values ('"&dmmc&"',"&dmfz&",'"&bz&"')"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " 添加成功!"))
	Call JsAlert("参考断面 "&dmmc&" 添加成功!","?action=ckfz")
end function
function ckfz_chgindb()
	dim dmmc, dmfz, bz
	dmmc="" : dmfz=0 : bz=""
	dmmc=trim(request("chgckdm"))
	dmfz=trim(request("chgdmfz"))
	bz=trim(request("chgbz"))
	if dmmc="" or dmfz="" or dmfz=0 or bz="" then response.write("数据不完整!请从正确入口进入!") : exit function
	strSql="update c_dmfz set dmfz="&dmfz&", bz='"&bz&"' where dmmc='"&dmmc&"'"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " 添加成功!"))
	Call JsAlert("参考断面 "&dmmc&" 更改成功!","?action=ckfz")
end function
function ckfz_delindb()
	dim dmmc
	dmmc=""
	dmmc=trim(request("delckdm"))
	if dmmc="" then response.write("数据不完整!请从正确入口进入!") : exit function
	strSql="delete from c_dmfz where dmmc='"&dmmc&"'"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " 添加成功!"))
	Call JsAlert("参考断面 "&dmmc&" 删除成功!","?action=ckfz")
end function
'----------------------------------------------------参考分值结束-----------------------------------------------------
%>