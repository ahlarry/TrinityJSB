<!--#include file="include/conn.asp"-->
<%
Call ChkAdminAble()
xjweb.Header()
dim action
action=request("action")
select case action
	case "depart"
		call common_dis("depart", "�û�����")
	case "depart_addindb"
		call common_add("depart", "�û�����")
	case "depart_delindb"
		call common_del("depart", "�û�����")

	case "sbcj"
		call common_dis("sbcj", "�豸����")
	case "sbcj_addindb"
		call common_add("sbcj", "�豸����")
	case "sbcj_delindb"
		call common_del("sbcj", "�豸����")

	case "dwmc"
		call common_dis("dwmc", "��λ����")
	case "dwmc_addindb"
		call common_add("dwmc", "��λ����")
	case "dwmc_delindb"
		call common_del("dwmc", "��λ����")
	
	case "mtjg"
		call common_dis("mtjg", "ģͷ�ṹ")
	case "mtjg_addindb"
		call common_add("mtjg", "ģͷ�ṹ")
	case "mtjg_delindb"
		call common_del("mtjg", "ģͷ�ṹ")
	
	case "dxjg"
		call common_dis("dxjg", "���ͽṹ")
	case "dxjg_addindb"
		call common_add("dxjg", "���ͽṹ")
	case "dxjg_delindb"
		call common_del("dxjg", "���ͽṹ")

	case "sxjg"
		call common_dis("sxjg", "ˮ��ṹ")
	case "sxjg_addindb"
		call common_add("sxjg", "ˮ��ṹ")
	case "sxjg_delindb"
		call common_del("sxjg", "ˮ��ṹ")

	case "dmmc"
		call common_dis("dmmc", "��������")
	case "dmmc_addindb"
		call common_add("dmmc", "��������")
	case "dmmc_delindb"
		call common_del("dmmc", "��������")

	case "mjcl"
		call common_dis("mjcl", "ģ�߲���")
	case "mjcl_addindb"
		call common_add("mjcl", "ģ�߲���")
	case "mjcl_delindb"
		call common_del("mjcl", "ģ�߲���")

	case "jcjxh"
		call common_dis("jcjxh", "�������ͺ�")
	case "jcjxh_addindb"
		call common_add("jcjxh", "�������ͺ�")
	case "jcjxh_delindb"
		call common_del("jcjxh", "�������ͺ�")

	case "rdogg"
		call common_dis("rdogg", "�ȵ�ż���")
	case "rdogg_addindb"
		call common_add("rdogg", "�ȵ�ż���")
	case "rdogg_delindb"
		call common_del("rdogg", "�ȵ�ż���")

	case "mtljcc"
		call common_dis("mtljcc", "ģͷ���ӳߴ�")
	case "mtljcc_addindb"
		call common_add("mtljcc", "ģͷ���ӳߴ�")
	case "mtljcc_delindb"
		call common_del("mtljcc", "ģͷ���ӳߴ�")

	case "lxrwlx"
		call common_dis("lxrwlx", "������������")
	case "lxrwlx_addindb"
		call common_add("lxrwlx", "������������")
	case "lxrwlx_delindb"
		call common_del("lxrwlx", "������������")

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
	Call TbTopoic("������ѡ����Ҫ��������Ŀ!")
end function

'-------------------------------------------------------���ú���------------------------------------------------------------------------------
rem ���ú���
function common_dis(item, stritem)		'��ʾ����   item Ϊ����, stritem-- Ϊ����   ��item=dwmc, stritem="��λ����"
	Call TbTopic(web_info(0) & "<font style=color:#ff0000;>"& stritem &"</font> -- ��ѯ����")
%>
	<table border=0 cellspacing=0 cellpadding=3 class=xtable width="60%">
		<form action="<%=request.servervariables("script_name")%>?action=<%=item%>_addindb" method="post" onsubmit="return chkaddinf();">
		<tr>
			<td class=rtd>���<%=stritem%></td>
			<td class=ltd><input type=text name="add<%=item%>"></td>
			<td class=ctd><input type=submit value=" ��� "></td>
		</tr>
		
		</form>
		<form action="<%=request.servervariables("script_name")%>?action=<%=item%>_delindb" method="post" onsubmit="return chkdelinf();">
		<tr>
			<td class=rtd>ɾ��<%=stritem%></td>
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
			<td class=ctd><input type=submit value=" ɾ�� "></td>
		</tr>
		</form>
	</table>
	<script language="javascript">
	function chkaddinf()
	{
		if(document.all.add<%=item%>.value==""){alert('������Ҫ��ӵ�<%=stritem%>!');document.all.add<%=item%>.focus();return false;}
	}
	function chkdelinf()
	{
		if(document.all.del<%=item%>.value=="")
			{alert('��ѡ��Ҫɾ����<%=stritem%>!');document.all.del<%=item%>.focus();return false;}
		else
			{return confirm('ȷ��Ҫɾ��<%=stritem%> '+document.all.del<%=item%>.value+' ��?');}
	}
	</script>
<%
end function

'---------------------------------------------�����⺯��-------------------------------------------------------------
function common_add(item, stritem)
	dim addinfo
	addinfo=trim(request("add"&item))
	if addinfo="" then response.write(prompt(stritem & "����Ϊ��!")) : exit function
	strSql="select "&item&" from [c_"&item&"] where "&item&"='"&addinfo&"'"
	set rs=xjweb.Exec(strSql, 1)
	if not(rs.eof or rs.bof) then Call JsAlert(""&stritem & " " & addinfo & " �Ѿ�����!","?action="&item&"") : Exit Function
	rs.close
	strSql="insert into c_"&item&" ("&item&") values ('"&addinfo&"')"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " ��ӳɹ�!"))
	Call JsAlert(""&stritem & " " & addinfo & " ��ӳɹ�!","?action="&item&"")
end function
'--------------------------------------------------�ӿ���ɾ������-----------------------------------------------------------
Function common_del(item, stritem)
	dim delinfo
	delinfo=trim(request("del"&item))
	if delinfo="" then response.write(prompt("��ѡ��Ҫɾ����"&stritem&"!")) : exit function
	strSql="delete from c_"&item&" where "&item&"='"&delinfo&"'"
	call xjweb.Exec(strSql, 0)
	'response.write(prompt(stritem & " " & delinfo & " ɾ���ɹ�!"))
	Call JsAlert(""&stritem & " " & delinfo & " ɾ���ɹ�!","?action="&item&"")
End Function
'----------------------------------------------------���ý���-----------------------------------------------------


'------------------------------------------------------��ֵ������ʼ--------------------------------------------------
function fzbl_dis()		'��ֵ����
	dim fzbl(15), fzblsm(15)
	strSql="select * from c_fzbl"
	set rs=xjweb.Exec(strSql,1)
	if rs.eof or rs.bof then rs.close : response.write ("ϵͳ����") : exit function
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
	fzblsm(0)="ģ�߷�ֵ"
	fzblsm(1)="ģ�߷�ֵ"
	fzblsm(2)="ģ�߷�ֵ"
	fzblsm(3)="ģ�߷�ֵ"
	fzblsm(4)="ģ�߷�ֵ"
	fzblsm(5)="ģ�߷�ֵ"
	fzblsm(6)="ģ�߷�ֵ"
	fzblsm(7)="ģ�߷�ֵ"
	fzblsm(8)="ģ�߷�ֵ"
	fzblsm(9)="ģ�߷�ֵ"
	fzblsm(10)="ģ�߷�ֵ"
	fzblsm(11)="ģ�߷�ֵ"
	fzblsm(12)="ģ�߷�ֵ"
	fzblsm(13)="ģ�߷�ֵ"
	fzblsm(14)="ģ�߷�ֵ"
	Call TbTopic(web_info(0) & " <font style=color:#ff0000;>��ֵ����(ϵ��)</font> -- ��ѯ����")
%>
	<table border=0 cellspacing=0 cellpadding=3 class=xtable width="60%">
		<form name="fzbl" action="<%=request.servervariables("script_name")%>?action=fzbl_indb" method="post" onsubmit="return chkinf();">
		<tr><td class=ctd colspan=3><b>��ֵ����(ϵ��)</b></td></tr>
		<tr>
			<td class=rtd width="40%">�ṹ����</td>
			<td class=ltd width="*"><input type=text name=jgbl size=6 value=<%=fzbl(0)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(0)%>"></td>
		</tr>
		<tr>
			<td class=rtd>��Ʊ���</td>
			<td class=ltd><input type=text name=sjbl size=6 value=<%=fzbl(1)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(1)%>"></td>
		</tr>
		<tr>
			<td class=rtd>��˱���</td>
			<td class=ltd><input type=text name=shbl size=6 value=<%=fzbl(2)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(2)%>"></td>
		</tr>
		<tr>
			<td class=rtd>���ı���</td>
			<td class=ltd><input type=text name=fgbl size=6 value=<%=fzbl(3)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(3)%>"></td>
		</tr>
		<tr>
			<td class=rtd>������˱���</td>
			<td class=ltd><input type=text name=fgshbl size=6 value=<%=fzbl(4)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(4)%>"></td>
		</tr>
		<tr>
			<td class=rtd>�������</td>
			<td class=ltd><input type=text name=fcbl size=6 value=<%=fzbl(5)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(5)%>"></td>
		</tr>
		<tr>
			<td class=rtd>�����ֵϵ��</td>
			<td class=ltd><input type=text name=gwfzxs size=6 value=<%=fzbl(6)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(6)%>"></td>
		</tr>
		<tr>
			<td class=rtd>��ʽ��ֵϵ��</td>
			<td class=ltd><input type=text name=msfzxs  size=6 value=<%=fzbl(7)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(7)%>"></td>
		</tr>
		<tr>
			<td class=rtd>�󹲼���ֵϵ��</td>
			<td class=ltd><input type=text name=hgjfzxs size=6 value=<%=fzbl(8)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(8)%>"></td>
		</tr>
		<tr>
			<td class=rtd>˫ɫ������ֵϵ��</td>
			<td class=ltd><input type=text name=ssgjfzxs size=6 value=<%=fzbl(9)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(9)%>"></td>
		</tr>
		<tr>
			<td class=rtd>��Ӳ������ֵϵ��</td>
			<td class=ltd><input type=text name=rygjfzxs size=6 value=<%=fzbl(10)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(10)%>"></td>
		</tr>
		<tr>
			<td class=rtd>BOM��ֵϵ��</td>
			<td class=ltd><input type=text name=bomfzxs  size=6 value=<%=fzbl(11)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(11)%>"></td>
		</tr>
		<tr>
			<td class=rtd>���Ե���ֵϵ��</td>
			<td class=ltd><input type=text name=tsdfzxs size=6 value=<%=fzbl(12)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(12)%>"></td>
		</tr>
		<tr>
			<td class=rtd>���Է�ֵϵ��</td>
			<td class=ltd><input type=text name=tsfzxs size=6 value=<%=fzbl(13)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(13)%>"></td>
		</tr>
		<tr>
			<td class=rtd>������Ϣ�����ֵϵ��</td>
			<td class=ltd><input type=text name=tsxxzlfzxs size=6 value=<%=fzbl(14)%>>%</td>
			<td class=ctd><img src="images/admin/help.gif" style="cursor:help;" alt="<%=fzblsm(14)%>"></td>
		</tr>
		<tr><td class=ctd colspan=3><input type=submit value=" ���� "></td></tr>
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
			call jschk("jgbl", "�ṹ����", fzbl(0))
			call jschk("sjbl", "��Ʊ���", fzbl(1))
			call jschk("shbl", "��˱���", fzbl(2))
			call jschk("fgbl", "���ı���", fzbl(3))
			call jschk("fgshbl", "������˱���", fzbl(4))
			call jschk("fcbl", "�������", fzbl(5))
			call jschk("gwfzxs", "�����ֵϵ��", fzbl(6))
			call jschk("msfzxs", "��ʽ��ֵϵ��", fzbl(7))
			call jschk("hgjfzxs", "�󹲼���ֵϵ��", fzbl(8))
			call jschk("ssgjfzxs", "˫ɫ������ֵϵ��", fzbl(9))
			call jschk("rygjfzxs", "��Ӳ������ֵϵ��", fzbl(10))
			call jschk("bomfzxs", "BOM��ֵϵ��", fzbl(11))
			call jschk("tsdfzxs", "���Ե���ֵϵ��", fzbl(12))
			call jschk("tsfzxs", "���Է�ֵϵ��", fzbl(13))
			call jschk("tsxxzlfzxs", "������Ϣ�����ֵϵ��", fzbl(14))			
		%>
	}
	</script>
<%
end function

function jschk(item, stritem, i)
%>
	if(isNaN(parseFloat(frm.<%=item%>.value))) 
		{alert('<%=stritem%>����Ϊ����!');frm.<%=item%>.value=<%=i%>;frm.<%=item%>.focus();return false;}
	frm.<%=item%>.value=parseFloat(frm.<%=item%>.value);
	if(<%=item%>>100 || <%=item%><0)	{alert('<%=stritem%>������ 0 �� 100 ֮��!');frm.<%=item%>.value=<%=i%>;frm.<%=item%>.focus();return false;}
<%
end function

'---------------------------------------------�����⺯��-------------------------------------------------------------
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
	Call JsAlert("��ֵ����(ϵ��) ���ĳɹ�!","?action=fzbl")
end function
'----------------------------------------------------��ֵ��������-----------------------------------------------------


'------------------------------------------------------�ο���ֵ��ʼ--------------------------------------------------
function ckfz_dis()		'�ο���ֵ
	Call TbTopic(web_info(0) & " <font style=color:#ff0000;>�ο���ֵ</font> -- ��ѯ����")
%>
	<table border=0 cellspacing=0 cellpadding=3 class=xtable width="80%">
		<form action="<%=request.servervariables("script_name")%>?action=ckfz_addindb" method="post" onsubmit="return chkaddinf();">
		<tr>
			<td class=rtd rowspan=3>��Ӳο�����</td>
			<td class=rtd>��������:</td>
			<td class=ltd><input type=text name="addckdm"></td>
			<td class=ctd rowspan=3><input type=submit value=" ��� "></td>
		</tr>
		<tr>
			<td class=rtd>�����ֵ:</td>
			<td class=ltd><input type=text size=10 name="adddmfz"></td>
		</tr>
		<tr>
			<td class=rtd>���ö���:</td>
			<td class=ltd><input type="text" size="60" name="addbz"></td></td>
		</tr>
		</form>
		<form action="<%=request.servervariables("script_name")%>?action=ckfz_chgindb" method="post" onsubmit="return chkchginf();">
		<tr>
			<td class=rtd rowspan=3>���Ĳο�����</td>
			<td class=rtd>��������:</td>
			<td class=ltd>
				<select name="chgckdm" onchange="if(this.selectedIndex==0){chgdmfz.value='';chgbz.value='';}else {chgdmfz.value=x_xcfz[this.selectedIndex-1]; chgbz.value=x_xcsm[this.selectedIndex-1];}">
				</select>
			</td>
			<td class=ctd rowspan=3><input type=submit value=" ���� "></td>
		</tr>
		<tr>
			<td class=rtd>�����ֵ:</td>
			<td class=ltd><input type=text size=10 name="chgdmfz"></td>
		</tr>
		<tr>
			<td class=rtd>���ö���:</td>
			<td class=ltd><input type="text" size="60" name="chgbz"></td>
		</tr>
		</form>
		<form action="<%=request.servervariables("script_name")%>?action=ckfz_delindb" method="post" onsubmit="return chkdelinf();">
		<tr>
			<td class=rtd rowspan=3>ɾ���ο�����</td>
			<td class=rtd>��������:</td>
			<td class=ltd>
				<select name="delckdm" onchange="if(this.selectedIndex==0){deldmfz.innerHTML='';delbz.innerHTML='';}else {deldmfz.innerHTML=x_xcfz[this.selectedIndex-1]; delbz.innerHTML=x_xcsm[this.selectedIndex-1];}">
				</select>
			</td>
			<td class=ctd rowspan=3><input type=submit value=" ɾ�� "></td>
		</tr>
		<tr>
			<td class=rtd>�����ֵ:</td>
			<td class=ltd><span id="deldmfz"></span>&nbsp;</td>
		</tr>
		<tr>
			<td class=rtd>���ö���:</td>
			<td class=ltd><span id="delbz"></span>&nbsp;</td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		//�Բο�ģ�߳�ʼ��
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
		if(document.all.addckdm.value=="")	{alert('�������Ʋ���Ϊ��!');document.all.addckdm.focus();return false;}	if(isNaN(parseFloat(document.all.adddmfz.value))){alert('�����ֵֻ��Ϊ����!');document.all.adddmfz.value='';document.all.adddmfz.focus();return false;}
		document.all.adddmfz.value=parseFloat(document.all.adddmfz.value)
		if(document.all.addbz.value=="")	{alert('���ö��治��Ϊ��!');document.all.addbz.focus();return false;}
	}

	function chkchginf()
	{
		if(document.all.chgckdm.value=="")	{alert('��ѡ��Ҫ���Ķ��������!');document.all.chgckdm.focus();return false;}
		if(isNaN(parseFloat(document.all.chgdmfz.value))){alert('�����ֵֻ��Ϊ����!');document.all.chgdmfz.value='';document.all.chgdmfz.focus();return false;}
		document.all.chgdmfz.value=parseFloat(document.all.chgdmfz.value)
		if(document.all.chgbz.value=="")	{alert('���ö��治��Ϊ��!');document.all.chgbz.focus();return false;}
	}
	function chkdelinf()
	{
		if(document.all.delckdm.value=="")	
			{alert('��ѡ��Ҫɾ�����������!');document.all.delckdm.focus();return false;}
		else
			{return confirm('ȷ��Ҫɾ���ο����� '+document.all.delckdm.value+' ��?')}
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
	if dmmc="" or dmfz="" or dmfz=0 or bz="" then response.write("���ݲ�����!�����ȷ��ڽ���!") : exit function
	strSql="select dmmc from c_dmfz where dmmc='"&dmmc&"'"
	set rs=xjweb.Exec(strSql, 1)
	if not(rs.eof or rs.bof) then Call JsAlert("�ο����� "&dmmc & " �Ѿ�����!","?action=ckfz") : Exit Function
	rs.close
	strSql="insert into c_dmfz (dmmc, dmfz, bz) values ('"&dmmc&"',"&dmfz&",'"&bz&"')"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " ��ӳɹ�!"))
	Call JsAlert("�ο����� "&dmmc&" ��ӳɹ�!","?action=ckfz")
end function
function ckfz_chgindb()
	dim dmmc, dmfz, bz
	dmmc="" : dmfz=0 : bz=""
	dmmc=trim(request("chgckdm"))
	dmfz=trim(request("chgdmfz"))
	bz=trim(request("chgbz"))
	if dmmc="" or dmfz="" or dmfz=0 or bz="" then response.write("���ݲ�����!�����ȷ��ڽ���!") : exit function
	strSql="update c_dmfz set dmfz="&dmfz&", bz='"&bz&"' where dmmc='"&dmmc&"'"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " ��ӳɹ�!"))
	Call JsAlert("�ο����� "&dmmc&" ���ĳɹ�!","?action=ckfz")
end function
function ckfz_delindb()
	dim dmmc
	dmmc=""
	dmmc=trim(request("delckdm"))
	if dmmc="" then response.write("���ݲ�����!�����ȷ��ڽ���!") : exit function
	strSql="delete from c_dmfz where dmmc='"&dmmc&"'"
	call xjweb.Exec(strSql, 0)
	'response.write (prompt(stritem & " " & addinfo & " ��ӳɹ�!"))
	Call JsAlert("�ο����� "&dmmc&" ɾ���ɹ�!","?action=ckfz")
end function
'----------------------------------------------------�ο���ֵ����-----------------------------------------------------
%>