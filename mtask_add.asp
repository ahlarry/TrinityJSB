<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtask_dbinf.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->
<%
'14:22 2007-1-6-������
Call ChkPageAble(3)
CurPage="������� �� ���������"
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
    <Td class=ctd height=300><%Call NewOrChange()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function NewOrChange()
	Dim s_lsh
	s_lsh=""
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call mtask_add() : Exit Function
	strSql="Select * from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("��ˮ�� " & s_lsh & " �����鲻����!","mtask_add.asp")
	Else
		Call mtask_cadd(Rs)
	End If
End Function

function mtask_cadd(rs)
	Dim idmmc
	idmmc=""
%>
<%Call TbTopic("�������������")%>
<table class=ktable cellspacing=0 cellpadding=3 width="98%">
  <form id=mtask_add name=mtask_add action=mtask_indb.asp?action=add method=post onSubmit='return checkinf();'>
    <tr>
      <th class=rtd height=25>��Ŀ����
        </td>
      <th colspan="2" class=ltd>��Ŀ����
        </td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>����ͬ��Ϣ��</b></td>
      <td colspan="2" class=ltd>��</td>
    </tr>
    <tr>
      <td class=rtd width="20%">������</td>
      <td colspan="2" class=ltd><input type=text name=ddh size=30 value=<%=rs("ddh")%>></td>
    </tr>
    <tr>
      <td class=rtd>��ˮ��</td>
      <td colspan="2" class=ltd><input type=text name=lsh size=30 onclick>
        &nbsp;�����ּ���ĸ���������ַ���</td>
    </tr>
    <tr>
      <td class=rtd>ģ��</td>
      <td colspan="2" class=ltd><input type=text name=mh size=30 value=<%=rs("mh")%>></td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <%If InStr(rs("dmmc"),"[")>0 Then
      		idmmc=Left(rs("dmmc"),Instr(rs("dmmc"),"[")-1)
      	Else
      		idmmc=rs("dmmc")
      	End If
      %>
      <td colspan="2" class=ltd><input type=text name=dmmc size=30 value="<%=idmmc%>">
        &nbsp;
        <select name="gfxl" onchange='changeselect(this.value);'>
          <option value="">��ѡ��</option>
          <%for i = 0 to ubound(c_gfxl)%>
          <option value='<%=c_gfxl(i) %>'><%=c_gfxl(i)%></option>
          <%next%>
        </select>
        &nbsp;
        <select name="gfdm" onchange='this.form.dmmc.value=this.form.dmmc.value+this.value;'>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>�ͻ�����</td>
      <td colspan="2" class=ltd><input type=text name=dwmc size=30 value=<%=rs("dwmc")%>>
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
      <td class=rtd>�豸����</td>
      <td colspan="2" class=ltd><input type=text name=sbcj size=30 value=<%=rs("sbcj")%>>
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
      <td class=rtd>�������ͺ�</td>
      <td colspan="2" class=ltd><input type=text name=jcjxh size=30 value=<%=rs("jcjxh")%>>
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
      <td class=rtd>ģ�߲���</td>
      <td colspan="2" class=ltd><input type=text name=mjcl size=30 value=<%=rs("mjcl")%>>
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
      <td class=rtd>ˮ��ͷ����</td>
      <td colspan="2" class=ltd><input type=text name=sjtsl size=30 value=<%=rs("sjtsl")%>></td>
    </tr>
    <tr>
      <td class=rtd>����ͷ����</td>
      <td colspan="2" class=ltd><input type=text name=qjtsl size=30 value=<%=rs("qjtsl")%>></td>
    </tr>
    <tr>
      <td class=rtd>ǣ���ٶ�</td>
      <td colspan="2" class=ltd><input type=text name=qysd size=10 value=<%=rs("qysd")%>>
        ��/��(m/min)</td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td colspan="2" class=ltd><select name=jcfx onchange=calmjfz();>
          <option value="/" <%if rs("jcfx")="/" then%> selected <%end if%>>&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <option value="��������" <%if rs("jcfx")="��������" then%> selected <%end if%>>��������</option>
          <option value="�ͻ�����" <%if rs("jcfx")="�ͻ�����" then%> selected <%end if%>>�ͻ�����</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>ǻ��</td>
      <td colspan="2" class=ltd><input type=text size=4 name=qs value=<%=rs("qs")%>>ǻ</td>
    </tr>
    <tr>
      <td class=rtd>����Ȱ�</td>
      <td colspan="2" class=ltd><select name="pjrb">
          <option value=true<%if rs("pjrb") then%> selected<%end if%>>��</option>
          <option value=false<%if not(rs("pjrb")) then%> selected<%end if%>>��</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>���Ȱ���Ϣ</td>
      <td colspan="2" class=ltd> ����:
        <select name="jrbxs">
          <option value="����"<%if rs("jrbxs")="����" then%> selected<%end if%>>����</option>
          <option value="����"<%if rs("jrbxs")="����" then%> selected<%end if%>>����</option>
        </select>
        &nbsp;
        ����:
        <select name="jrbcl">
          <option value="����" <%if rs("jrbcl")="����" then%> selected <%end if%>>����</option>
          <option value="��ĸ"<%if rs("jrbcl")="��ĸ" then%> selected <%end if%>>��ĸ</option>
        </select>
        &nbsp;
        ����˵��:
        <input type=text name=jrbxx size=40 value=<%=rs("jrbxx")%>></td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>����ֵ��Ϣ��</b></td>
      <td colspan="2" class=ltd>(�������������ʱ������ѡ��)</td>
    </tr>
    <tr>
      <td class=rtd>�ο�����</td>
      <td colspan="2" class=ltd alt="1.ѡ��ο�����:���ģ�ߵĴ�ŷ�ֵ;<br>2.ѡ����ϵ��:ȷ��ģ�ߵľ����ֵ;<br>3.���ݲ�ͬ��ģ�����ѡ��:ȷ��ģ�ߵ����շ�ֵ."><select name="ckdm" onChange="if(this.selectedIndex==0) xcsm.innerHTML='';else xcsm.innerHTML=' ��ֵ:'+x_xcfz[this.selectedIndex-1] + '  �ʺ���:' + x_xcsm[this.selectedIndex-1];calmjfz();">
        </select>
        &nbsp;&nbsp; <span id=xcsm></span></td>
    </tr>
    <tr>
      <td class=rtd>�������</td>
      <td class=ltd alt="1.ѡ��ο�����:���ģ�ߵĻ�������;<br>2.ѡ����ϵ��:ȷ��ģ�ߵľ��嶨��;<br>3.���ݲ�ͬ��ģ�����ѡ��:ȷ��ģ�ߵ����ն���."><select name="dedm" onChange="if(this.selectedIndex==0) jcde.innerHTML='';else jcde.innerHTML=' ����:'+x_defz[this.selectedIndex-1];document.all.defz.value=x_defz[this.selectedIndex-1];">
        </select>
        &nbsp;&nbsp; <span id=jcde></span></td>
    </tr>        
    <tr>
      <td class=rtd>����ϵ��</td>
      <td colspan="2" class=ltd><input name=fzxs type=text onchange="calmjfz()" size=5 value=1></td>
    </tr>
    <tr>
      <td class=rtd>����</td>
      <td  class=ltd>
        <input type="checkbox" name="gjfs1" class="radio" value="1" onclick="calmjfz();" <%If NullToNum(Rs("ssgj"))<>0 Then%> checked <%End If%> />
        ˫ɫ����<input name=ssgjf type=text onchange="calmjfz()" value=<%=Rs("ssgj")%> size=5 <%If NullToNum(Rs("ssgj"))=0 Then%> style="display:none" <%End If%>>
        <input type="checkbox" name="gjfs2" class="radio" value="1" onclick="calmjfz();" <%If NullToNum(Rs("qbfgj"))<>0 Then%> checked <%End If%> />
        ȫ��������<input name=qbfgjf type=text onchange="calmjfz()" value=<%=Rs("qbfgj")%> size=5  <%If NullToNum(Rs("qbfgj"))=0 Then%> style="display:none" <%End If%>>
        <input type="checkbox" name="gjfs3" class="radio" value="1" onclick="document.mtask_add.gjfs4.checked=false;calmjfz();" <%If NullToNum(Rs("qgj"))<>0 Then%> checked <%End If%> />
        ��Ӳǰ����<input name=qgjf type=text onchange="calmjfz()" value=<%=Rs("qgj")%> size=5  <%If NullToNum(Rs("qgj"))=0 Then%> style="display:none" <%End If%>>
        <input type="checkbox" name="gjfs4" class="radio" value="1" onclick="document.mtask_add.gjfs3.checked=false;calmjfz();" <%If NullToNum(Rs("hgj"))<>0 Then%> checked <%End If%> />
        ��Ӳ�󹲼� <input name=hgjf type=text onchange="calmjfz()" value=<%=Rs("hgj")%> size=5  <%If NullToNum(Rs("hgj"))=0 Then%> style="display:none" <%End If%>> </td>
    </tr>
    <tr>
      <td class=rtd>ģ�߷�ֵ</td>
      <td colspan="2" class=ltd> ģ���ܷ�:<span id=span_mjzf style="font-weight:bold;">0</span>��&nbsp;&nbsp;&nbsp;&nbsp; <span id=span_gjzf style="font-weight:bold;">0</span> <br>
        BOM�ܷ�:<span id=span_bomzf style="font-weight:bold;">0</span>��<br>
        �����ֲ��ܷ�:<span id=span_tsdzf style="font-weight:bold;">0</span>��<br>
        �����ܷ�:<span id=span_tszf style="font-weight:bold;">0</span>��<br>
        ������Ϣ�����ܷ�:<span id=span_tsxxzlzf style="font-weight:bold;">0</span>��<br></td>
    </tr>
    <input type=hidden name=mjzf value=0>
    <input type=hidden name=gjzf value=0>
    <input type=hidden name=bomzf value=0>
    <input type=hidden name=tsdzf value=0>
    <input type=hidden name=tszf value=0>
    <input type=hidden name=tsxxzlzf value=0>
    <input type=hidden name=bgw value=false>
    <input type=hidden name=bms value=false>
    <input type=hidden name=bhgj value=false>
    <input type=hidden name=bssgj value=false>
    <input type=hidden name=brygj value=false>
    <input type=hidden name=defz value=0>
    <tr>
      <td class=rtd rowspan="2">��ֵ����</td>
      <td class=ltd>ģͷ����:
        <input type=text name=mtbl size=4 value="40" onchange=blchange();>
        %&nbsp;&nbsp;&nbsp;���ͱ���:
        <input type=text name=dxbl size=4  value="60" disabled>
        %</td>
    </tr>
    <tr>
        <td class=ltd>ģͷ�ṹ:
        <input type=text name=mtjgbl size=4 >
        %&nbsp;&nbsp;&nbsp;���ͽṹ:
        <input type=text name=dxjgbl size=4 > %</td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>��ģ����Ϣ��</b></td>
      <td colspan="2" class=ltd>��</td>
    </tr>
    <tr>
      <td class=rtd>ģ����Ϣ</td>
      <td colspan="2" class=ltd><select name=mjxx onchange='chkmjxx(this);'>
          <option value="ȫ��" <%if rs("mjxx")="ȫ��" then%> selected <%end if%>>ȫ��</option>
          <option value="ģͷ" <%if rs("mjxx")="ģͷ" then%> selected <%end if%>>ģͷ</option>
          <option value="����" <%if rs("mjxx")="����" then%> selected <%end if%>>����</option>
        </select>&nbsp;&nbsp;&nbsp;
        ģͷ<select name=mtrw id="mtrw">
          <option value="" selected></option>
          <option value="���" <%if rs("mtrw")="���" then%> selected <%end if%>>���</option>
          <option value="����" <%if rs("mtrw")="����" then%> selected <%end if%>>����</option>
          <option value="����" <%if rs("mtrw")="����" then%> selected <%end if%>>����</option>
        </select>&nbsp;&nbsp;&nbsp;
        ����<select name=dxrw id="dxrw">
          <option value="" selected></option>        
          <option value="���" <%if rs("dxrw")="���" then%> selected <%end if%>>���</option>
          <option value="����" <%if rs("dxrw")="����" then%> selected <%end if%>>����</option>
          <option value="����" <%if rs("dxrw")="����" then%> selected <%end if%>>����</option>
        </select></td>        
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td colspan="2" class=ltd><select name=rwlr>
          <option value="���" <%if rs("rwlr")="���" then%> selected <%end if%>>���</option>
          <option value="����" <%if rs("rwlr")="����" then%> selected <%end if%>>����</option>
          <option value="����" <%if rs("rwlr")="����" then%> selected <%end if%>>����</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>���ڵ���</td>
      <td colspan="2" class=ltd><select name="cnts" style="width:51px;" onchange='ExcTslb(this);'>
          <option value=true <%if rs("cnts") then%> selected <%end if%>>��</option>
          <option value=false <%if not(rs("cnts")) then%> selected <%end if%>>��</option>
        </select></td>
    </tr>
    <tr id=trbeit>
      <td class=rtd>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
      <td colspan="2" class=ltd><select name=beit style="width:51px;">
          <option value=true <%if rs("beit") then%> selected <%end if%>>��</option>
          <option value=false <%if not(rs("beit")) then%> selected <%end if%>>��</option>
        </select></td>
    </tr>
    <tr id=trtslb>
      <td class=rtd>�������</td>
      <td colspan="2" class=ltd alt="1.ѡ���Ͳ����:ȷ��ģ�ߵ����ɵ��Դ�����<br>2.ʵ�ʴ������������Ĳ�˵��ģ�߽ṹ�����Է�������ȷ��."><select name="tslb" style="width:51px;" onChange="if(this.selectedIndex==0) xcts.innerHTML='';else xcts.innerHTML=' ����Դ���:'+z_xccs[this.selectedIndex-1] + '  ������Χ:' + z_xcfw[this.selectedIndex-1] + '  ������:' + z_xcbz[this.selectedIndex-1];">
        </select>
        &nbsp;&nbsp; <span id=xcts></span></td>
    </tr>
    <tr>
      <td class=rtd>�Ͳıں�</td>
      <td colspan="2" class=ltd><input type=text name=xcbh size=14 value=<%=Rs("xcbh")%>></td>
    </tr>
    <tr id=trdxqg>
      <td class=rtd>�����и�</td>
      <td colspan="2" class=ltd><select name="dxqg">
          <option value="���ϸ�"<%if rs("dxqg")="���ϸ�" then%> selected<%end if%>>���ϸ�</option>
          <option value="����ϸ�"<%if rs("dxqg")="����ϸ�" then%> selected<%end if%>>����ϸ�</option>
          <option value="����ϸ�"<%if rs("dxqg")="����ϸ�" then%> selected<%end if%>>����ϸ�</option>
          <option value="����һ���и�"<%if rs("dxqg")="����һ���и�" then%> selected<%end if%>>����һ���и�</option>
        </select></td>
    </tr>
    <tr id=trdxjg>
      <td class=rtd>���ͽṹ</td>
      <td colspan="2" class=ltd><input type=text name=dxjg size=30 value=<%=rs("dxjg")%>>
        &nbsp;
        <select onchange='this.form.dxjg.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_dxjg)%>
          <option value='<%=c_dxjg(i)%>'><%=c_dxjg(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr id=trsxjg>
      <td class=rtd>ˮ��ṹ</td>
      <td colspan="2" class=ltd><input type=text name=sxjg size=30 value=<%=rs("sxjg")%>>
        &nbsp;
        <select onchange='this.form.sxjg.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_sxjg)%>
          <option value='<%=c_sxjg(i)%>'><%=c_sxjg(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>�ȵ�ż���</td>
      <td colspan="2" class=ltd><input type=text name=rdogg size=30 value="<%=rs("rdogg")%>">
        &nbsp;
        <select onchange='this.form.rdogg.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_rdogg)%>
          <option value='<%=c_rdogg(i)%>'><%=c_rdogg(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>ģͷ���ӳߴ�</td>
      <td colspan="2" class=ltd><input type=text name=mtljcc size=40 value="<%=rs("mtljcc")%>">
        &nbsp;
        <select onchange='this.form.mtljcc.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_mtljcc)%>
          <option value='<%=c_mtljcc(i)%>'><%=c_mtljcc(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>�������ӳߴ�</td>
      <td colspan="2" class=ltd><input type=text name=gjljcc size=40 value="<%=rs("gjljcc")%>"></td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>��������Ϣ��</b></td>
      <td colspan="2" class=ltd>��</td>
    </tr>
    <tr>
      <td class=rtd>��ע</td>
      <td colspan="2" class=ltd><textarea name="bz" cols="75" rows="7"><%=rs("bz")%></textarea></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ���ʼʱ��</td>
      <td colspan="2" class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jhkssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ��ṹ����ʱ��</td>
      <td colspan="2" class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jgjssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ�ȫ�׽���ʱ��</td>
      <td colspan="2" class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jhjssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ṹ�鳤</td>
      <td colspan="2" class=ltd><select name="jgzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%if rs("jgzz")=c_allzz(i) then%> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>����鳤</td>
      <td colspan="2" class=ltd><select name="sjzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%if rs("sjzz")=c_allzz(i) then%> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td colspan="2" class=ltd><select name="jsdb" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'<%if rs("jsdb")=c_allzy(i) then%> selected<%end if%>><%=c_allzy(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=ctd colspan=3><input type=submit value=" �� ȷ �� �� "></td>
    </tr>
  </form>
</table>
<%
	Dim TmpTslb,TmpCkdm,TmpDmde
	TmpTslb=Rs("tslb")
	TmpCkdm=Rs("ckdm")
	TmpDmde=Rs("dedm")
	call mtask_js(TmpTslb,TmpCkdm,TmpDmde)
end function		'mtask_cadd()


Function mtask_add()
%>
<%Call TbTopic("���������")%>
<table class=ktable cellspacing=0 cellpadding=3 width="95%">
  <form id=mtask_add name=mtask_add action=mtask_indb.asp?action=add method=post onSubmit='return checkinf();'>
    <tr>
      <th class=ctd height=25>��Ŀ����
        </td>
      <th class=ctd>��Ŀ����
        </td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>����ͬ��Ϣ��</b></td>
      <td class=ltd>��</td>
    </tr>
    <tr>
      <td class=rtd width="20%">������</td>
      <td class=ltd><input type=text name=ddh size=30></td>
    </tr>
    <tr>
      <td class=rtd>��ˮ��</td>
      <td class=ltd><input type=text name=lsh size=30>
        &nbsp;�����ּ���ĸ���������ַ���</td>
    </tr>
    <tr>
      <td class=rtd>ģ��</td>
      <td class=ltd><input type=text name=mh size=30></td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td class=ltd><input type=text name=dmmc size=30>
        &nbsp;
        <select name="gfxl" onchange='changeselect(this.value);'>
          <option value="">��ѡ��</option>
          <%for i = 0 to ubound(c_gfxl)%>
          <option value='<%=c_gfxl(i) %>'><%=c_gfxl(i)%></option>
          <%next%>
        </select>
        &nbsp;
        <select name="gfdm" onchange='this.form.dmmc.value=this.form.dmmc.value+this.value;'>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>�ͻ�����</td>
      <td class=ltd><input type=text name=dwmc size=30>
        <span style="width:18px;border:0px solid red;">
        <select name="dwmc_c" style="margin-left:-200px;width:218px;" onChange='this.form.dwmc.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_dwmc)%>
          <option value='<%=c_dwmc(i)%>'><%=c_dwmc(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>�豸����</td>
      <td class=ltd><input type=text name=sbcj size=30>
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
      <td class=rtd>�������ͺ�</td>
      <td class=ltd><input type=text name=jcjxh size=30>
        <span style="width:18px;border:0px solid red;">
        <select  style="margin-left:-200px;width:218px;" onchange='this.form.jcjxh.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_jcjxh)%>
          <option value='<%=c_jcjxh(i)%>'><%=c_jcjxh(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>ģ�߲���</td>
      <td class=ltd><input type=text name=mjcl size=30>
        <span style="width:18px;border:0px solid red;">
        <select  style="margin-left:-200px;width:218px;" onchange='this.form.mjcl.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_mjcl)%>
          <option value='<%=c_mjcl(i)%>'><%=c_mjcl(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>ˮ��ͷ����</td>
      <td class=ltd><input type=text name=sjtsl size=30></td>
    </tr>
    <tr>
      <td class=rtd>����ͷ����</td>
      <td class=ltd><input type=text name=qjtsl size=30></td>
    </tr>
    <tr>
      <td class=rtd>ǣ���ٶ�</td>
      <td class=ltd><input type=text name=qysd size=10>
        ��/��(m/min)</td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td class=ltd><select name=jcfx onchange=calmjfz();>
          <option value="/">&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <option value="��������" selected>��������</option>
          <option value="�ͻ�����">�ͻ�����</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>ǻ��</td>
      <td class=ltd><input type=text size=4 name=qs>ǻ</td>
    </tr>
    <tr>
      <td class=rtd>����Ȱ�</td>
      <td class=ltd><select name="pjrb">
          <option value=true>��</option>
          <option value=false selected>��</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>���Ȱ���Ϣ</td>
      <td class=ltd> ����:
        <select name="jrbxs">
          <option value="����">����</option>
          <option value="����">����</option>
        </select>
        &nbsp;
        ����:
        <select name="jrbcl">
          <option value="����">����</option>
          <option value="��ĸ">��ĸ</option>
        </select>
        &nbsp;
        ����˵��:
        <input type=text name=jrbxx size=40></td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>����ֵ��Ϣ��</b></td>
      <td class=ltd>��</td>
    </tr>
    <tr>
      <td class=rtd>�ο�����</td>
      <td class=ltd alt="1.ѡ��ο�����:���ģ�ߵĴ�ŷ�ֵ;<br>2.ѡ����ϵ��:ȷ��ģ�ߵľ����ֵ;<br>3.���ݲ�ͬ��ģ�����ѡ��:ȷ��ģ�ߵ����շ�ֵ."><select name="ckdm" onChange="if(this.selectedIndex==0) xcsm.innerHTML='';else xcsm.innerHTML=' ��ֵ:'+x_xcfz[this.selectedIndex-1] + '  �ʺ���:' + x_xcsm[this.selectedIndex-1];calmjfz();">
        </select>
        &nbsp;&nbsp; <span id=xcsm></span></td>
    </tr>
    <tr>
      <td class=rtd>�������</td>
      <td class=ltd alt="1.ѡ��ο�����:���ģ�ߵĻ�������;<br>2.ѡ����ϵ��:ȷ��ģ�ߵľ��嶨��;<br>3.���ݲ�ͬ��ģ�����ѡ��:ȷ��ģ�ߵ����ն���."><select name="dedm" onChange="if(this.selectedIndex==0) jcde.innerHTML='';else jcde.innerHTML=' ����:'+x_defz[this.selectedIndex-1];document.all.defz.value=x_defz[this.selectedIndex-1];">
        </select>
        &nbsp;&nbsp; <span id=jcde></span></td>
    </tr>    
    <tr>
      <td class=rtd>����ϵ��</td>
      <td class=ltd><input name=fzxs type=text onchange="calmjfz()" size=5 value=1></td>
    </tr>
    <tr>
      <td class=rtd>����</td>
      <td  class=ltd>
        <input type="checkbox" name="gjfs1" class="radio" value="1" onclick="calmjfz();" />
        ˫ɫ����<input name=ssgjf type=text onchange="calmjfz()" value="0" size=5 style="display:none">
        <input type="checkbox" name="gjfs2" class="radio" value="1" onclick="calmjfz();" />
        ȫ��������<input name=qbfgjf type=text onchange="calmjfz()" value="0" size=5 style="display:none">
        <input type="checkbox" name="gjfs3" class="radio" value="1" onclick="document.mtask_add.gjfs4.checked=false;calmjfz();">
        ��Ӳǰ����<input name=qgjf type=text onchange="calmjfz()" value="0" size=5 style="display:none">
        <input type="checkbox" name="gjfs4" class="radio" value="1" onclick="document.mtask_add.gjfs3.checked=false;calmjfz();">
        ��Ӳ�󹲼�<input name=hgjf type=text onchange="calmjfz()" value="0" size=5 style="display:none"> </td>
    </tr>
    <tr>
      <td class=rtd>ģ�߷�ֵ</td>
      <td class=ltd> ģ���ܷ�:<span id=span_mjzf style="font-weight:bold;">0</span>��&nbsp;&nbsp;&nbsp;&nbsp; <span id=span_gjzf style="font-weight:bold;"></span> <br>
        BOM�ܷ�:<span id=span_bomzf style="font-weight:bold;">0</span>��<br>
        �����ֲ��ܷ�:<span id=span_tsdzf style="font-weight:bold;">0</span>��<br>
        �����ܷ�:<span id=span_tszf style="font-weight:bold;">0</span>��<br>
        ������Ϣ�����ܷ�:<span id=span_tsxxzlzf style="font-weight:bold;">0</span>��<br></td>
    </tr>
    <input type=hidden name=mjzf value=0>
    <input type=hidden name=gjzf value=0>
    <input type=hidden name=bomzf value=0>
    <input type=hidden name=tsdzf value=0>
    <input type=hidden name=tszf value=0>
    <input type=hidden name=tsxxzlzf value=0>
    <input type=hidden name=bssgj value=false>
    <input type=hidden name=bqbfgj value=false>
    <input type=hidden name=bryqgj value=false>
    <input type=hidden name=bryhgj value=false>
    <input type=hidden name=defz value=0>
    
    <tr>
      <td class=rtd rowspan="2">��ֵ����</td>
      <td class=ltd>ģͷ����:
        <input type=text name=mtbl size=4 value="40" onchange=blchange();>
        %&nbsp;&nbsp;&nbsp;���ͱ���:
        <input type=text name=dxbl size=4  value="60" disabled>
        %</td>
    </tr>
    <tr>
        <td class=ltd>ģͷ�ṹ:
        <input type=text name=mtjgbl size=4 >
        %&nbsp;&nbsp;&nbsp;���ͽṹ:
        <input type=text name=dxjgbl size=4 > %</td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>��ģ����Ϣ��</b></td>
      <td class=ltd>��</td>
    </tr>
    <tr>
      <td class=rtd>ģ����Ϣ</td>
      <td class=ltd><select name=mjxx onchange='chkmjxx(this);'>
          <option value="ȫ��" selected>ȫ��</option>
          <option value="ģͷ">ģͷ</option>
          <option value="����">����</option>
        </select>&nbsp;&nbsp;&nbsp;
      ģͷ<select name=mtrw id="mtrw">
          <option value=""></option>
          <option value="���">���</option>
          <option value="����">����</option>
          <option value="����">����</option>
        </select>&nbsp;&nbsp;&nbsp;
      ����<select name=dxrw id="dxrw">
          <option value=""></option>
          <option value="���">���</option>
          <option value="����">����</option>
          <option value="����">����</option>
        </select></td>
    </tr>     
    <tr>
      <td class=rtd>��������</td>
      <td class=ltd><select name=rwlr>
          <option value="���" selected>���</option>
          <option value="����">����</option>
          <option value="����">����</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>���ڵ���</td>
      <td class=ltd><select name=cnts style="width:51px;" onchange='ExcTslb(this);'>
          <option value=true selected>��</option>
          <option value=false>��</option>
        </select></td>
    </tr>
    <tr id=trbeit>
      <td class=rtd>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
      <td class=ltd><select name=beit style="width:51px;">
          <option value=true selected>��</option>
          <option value=false>��</option>
        </select></td>
    </tr>
    <tr id=trtslb>
      <td class=rtd>�������</td>
      <td class=ltd alt="1.ѡ���Ͳ����:ȷ��ģ�ߵ����ɵ��Դ�����<br>2.ʵ�ʴ������������Ĳ�˵��ģ�߽ṹ�����Է�������ȷ��."><select name="tslb" style="width:51px;" onChange="if(this.selectedIndex==0) xcts.innerHTML='';else xcts.innerHTML=' ����Դ���:'+z_xccs[this.selectedIndex-1] + ' - ' + z_xcfw[this.selectedIndex-1] + '  ������:' + z_xcbz[this.selectedIndex-1];">
        </select>
        &nbsp;&nbsp; <span id=xcts></span></td>
    </tr>
    <tr>
      <td class=rtd>�Ͳıں�</td>
      <td class=ltd><input type=text name=xcbh size=14></td>
    </tr>
    <tr  id=trdxqg>
      <td class=rtd>�����и�</td>
      <td class=ltd><select name="dxqg">
          <option value="���ϸ�" selected>���ϸ�</option>
          <option value="����ϸ�">����ϸ�</option>
          <option value="����ϸ�">����ϸ�</option>
          <option value="����һ���и�">����һ���и�</option>
        </select></td>
    </tr>
    <tr id=trdxjg>
      <td class=rtd>���ͽṹ</td>
      <td class=ltd><input type=text name=dxjg size=30>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onchange='this.form.dxjg.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_dxjg)%>
          <option value='<%=c_dxjg(i)%>'><%=c_dxjg(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr id=trsxjg>
      <td class=rtd>ˮ��ṹ</td>
      <td class=ltd><input type=text name=sxjg size=30>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onchange='this.form.sxjg.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_sxjg)%>
          <option value='<%=c_sxjg(i)%>'><%=c_sxjg(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>�ȵ�ż���</td>
      <td class=ltd><input type=text name=rdogg size=30>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onchange='this.form.rdogg.value=this.value;'>
          <option></option>
          "
          <%for i = 0 to ubound(c_rdogg)%>
          <option value='<%=c_rdogg(i)%>'><%=c_rdogg(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>ģͷ���ӳߴ�</td>
      <td class=ltd><input type=text name=mtljcc size=30>
        <span style="width:18px;border:0px solid red;">
        <select style="margin-left:-200px;width:218px;" onchange='this.form.mtljcc.value=this.value;'>
          <option></option>
          <%for i = 0 to ubound(c_mtljcc)%>
          <option value='<%=c_mtljcc(i)%>'><%=c_mtljcc(i)%></option>
          <%next%>
        </select>
        </span></td>
    </tr>
    <tr>
      <td class=rtd>�������ӳߴ�</td>
      <td class=ltd><input type=text name=gjljcc size=30 value="/"></td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>��������Ϣ��</b></td>
      <td class=ltd>��</td>
    </tr>
    <tr>
      <td class=rtd>��ע</td>
      <td class=ltd><textarea name="bz" cols="75" rows="7"></textarea></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ���ʼʱ��</td>
      <td class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jhkssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ��ṹ����ʱ��</td>
      <td class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jgjssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ�ȫ�׽���ʱ��</td>
      <td class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jhjssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ṹ�鳤</td>
      <td class=ltd><select name="jgzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>����鳤</td>
      <td class=ltd><select name="sjzz"  style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td class=ltd><select name="jsdb" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=ctd colspan=2><input type=submit value=" �� ȷ �� �� "></td>
    </tr>
  </form>
</table>
<%
	call mtask_js("","","")
end function		'mtask_add()
%>
<%
function mtask_js(TslbOv,CkdmOV,DmdeOV)
'����ΪJS����%>
<script language="javascript">
//�Բο�ģ�߳�ʼ��
	var x_xcmc = new Array();
	var x_xcfz = new Array();
	var x_xcsm = new Array();

<%
	set rs=xjweb.exec("select * from c_dmfz order by dmmc",1)
	i=0
	do while not rs.eof
%>
		x_xcmc[<%=i%>]="<%=rs("dmmc")%>";
		x_xcfz[<%=i%>]="<%=rs("dmfz")%>";
		x_xcsm[<%=i%>]="<%=rs("bz")%>";
<%
		i = i + 1
		rs.movenext
	loop
	rs.close
%>
	for(var i=1; i<x_xcmc.length + 1; i++)
	{
		document.all.ckdm[i] = new Option(x_xcmc[i-1],x_xcmc[i-1]);
		if(document.all.ckdm.options[i].value=="<%=CkdmOV%>")
 			document.all.ckdm.options[i].selected=true; 
	}
	calmjfz();
		
//�Զ�������ʼ��
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
		
//�Ե�������ʼ��
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
//����ģ�߷�ֵ,ͬʱ��ʾ���ز�
function calmjfz()
{
	//��ֵϵ����ʼ��(��ʽʹ��ʱ�ӿ��ж�ȡ)
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

	var ttmjfz=0;		//ģ���ܷ�ֵ
	var ttgjfz=0;		//������ֵ
	var tmpobj;
	var tmpstr;
	document.all.span_gjzf.innerHTML="";
	if(isNaN(parseFloat(document.all.mtjgbl.value))) document.all.mtjgbl.value=Math.round(mtjgbl*100);
	if(isNaN(parseFloat(document.all.dxjgbl.value))) document.all.dxjgbl.value=Math.round(dxjgbl*100);

	//���������ֵ
	var str=document.all;
	//�ɲο������ó�ʼ��ֵ
	if((str.ckdm.selectedIndex-1)>=0) ttmjfz=x_xcfz[str.ckdm.selectedIndex-1];

	//����ȷ��;
	var issgjf=str.ssgjf.value*1;
	var iqbfgjf=str.qbfgjf.value*1;
	var iqgjf=str.qgjf.value*1;
	var ihgjf=str.hgjf.value*1;
	if (str.gjfs1.checked)	//˫ɫ����
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
		str.span_gjzf.innerHTML="����: " + Math.round(ttgjfz) + "��"
	}
	else
	{
		str.ssgjf.value=0;
		str.ssgjf.style.display="none";
	}
	if (str.gjfs2.checked)	//ȫ��������
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
		str.span_gjzf.innerHTML="����: " + Math.round(ttgjfz) + "��"
	}
	else
	{
		str.qbfgjf.value=0;
		str.qbfgjf.style.display="none";
	}
	if (str.gjfs3.checked)	//��Ӳǰ����
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
		str.span_gjzf.innerHTML="����: " + Math.round(ttgjfz) + "��"
	}
	else
	{
		str.qgjf.value=0;
		str.qgjf.style.display="none";
	}
	if (str.gjfs4.checked)	//��Ӳ�󹲼�
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
		str.span_gjzf.innerHTML="����: " + Math.round(ttgjfz) + "��"
	}
	else
	{
		str.hgjf.value=0;
		str.hgjf.style.display="none";
	}

	//ǻ��ȷ��
	//ttmjfz=ttmjfz*(Math.sqrt(str.qs.value));

	//����ϵ��ȷ��
	ttmjfz=ttmjfz*(str.fzxs.value);

	//ģ����Ϣ��ģͷ�����ͣ�ȷ��	
	switch (str.mjxx.value)
	{
		case "ģͷ" :
			ttmjfz=ttmjfz*0.4;
			break;
		case "����" :
			ttmjfz=ttmjfz*0.6;
			break;
		default:
			break;
	}

	//ģ���ܷ�
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

//��ֵ����
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

//ģ����Ϣ
function chkmjxx(ftemp)
{
	switch (ftemp.value)
	{
		case "ģͷ" :
			document.all.trdxqg.style.display="none";
			document.all.trdxjg.style.display="none";
			document.all.trsxjg.style.display="none";
			document.all.dxqg.value="";
			document.all.dxjg.value="/";
			document.all.sxjg.value="/";
			document.all.mtbl.value="100";
			document.all.dxbl.value="0";
                   document.getElementById("dxrw").value="";      
			document.getElementById("mtrw").disabled=false;
			document.getElementById("dxrw").disabled=true;
			break;
		case "����" :
			document.all.trdxqg.style.display="";
			document.all.trdxjg.style.display="";
			document.all.trsxjg.style.display="";
			document.getElementById("mtrw").value="";
			document.all.mtbl.value="0";
			document.all.dxbl.value="100";
			document.getElementById("mtrw").disabled=true;
			document.getElementById("dxrw").disabled=false;
			break;
		default:
			document.all.trdxqg.style.display="";
			document.all.trdxjg.style.display="";
			document.all.trsxjg.style.display="";
			document.all.mtbl.value="40";
			document.all.dxbl.value="60";
			document.getElementById("mtrw").disabled=false;
			document.getElementById("dxrw").disabled=false;		
			break;
	}
	calmjfz()
}
//Ĭ��״̬�±��������˵�����ʾ
document.all.trbeit.style.display="none";
//���ڵ��Թ�������
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

//�������ƹ�������
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
	document.all.gfdm.options[document.all.gfdm.length] = new Option("��ѡ��","");
	for (j=0;j<igfdm.length;j++){
		if(igfdm[j][0]==selvalue){
			document.all.gfdm.options[document.all.gfdm.length] = new Option(igfdm[j][1],igfdm[j][1]);
		}
	}
}
document.all.gfdm.options[document.all.gfdm.length] = new Option("��ѡ��","");
chkmjxx(document.all.mjxx);
</script>
<%
end function	'mtask_js()
%>
