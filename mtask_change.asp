<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtask_dbinf.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->
<%
'10:52 2007-1-25-������
Call ChkPageAble(3)
CurPage="������� �� ����������"
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
	If s_lsh="" Then Call TbTopic("������Ҫ���ĵ����������ˮ��!") : Exit Function

	DO while CC<2
	strSql="select * from [mtask] where [lsh]='"&s_lsh&"'"
	set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		s_lsh="C"&s_lsh
		CC=CC+1
		If CC=2 Then
			Call JsAlert("��ˮ��Ϊ �� " & right(s_lsh,len(s_lsh)-1) & " �� �������鲻����!", "mtask_change.asp")
		End If
	Else
		CC=3
		If Not(IsNull(Rs("sjjssj"))) Then
			s_time=datediff("d", rs("sjjssj"), now)
			if s_time > 5 then
	'			Call JsAlert("��ˮ��Ϊ �� " & s_lsh & " �� ���������Ѿ����"& s_time &" ��,�ѹ��ɱ༭����!", "mtask_change.asp")
				Call BzChan(Rs)
			else %>
<Script language="javascript">
				alert("��<%=s_lsh%>�������������,<%=5-s_time%>��֮���Կɸ��ġ�");
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
<%Call TbTopic("������ˮ�� <font style=color:#0000FF>" & Rs("lsh") & "</font> ��������")%>
<%If ChkAble(4) Then Response.Write "<a href=mtask_zzchange.asp?s_lsh="&Rs("lsh")&">�鳤Ȩ��</a><br>"%>
<table class=xtable cellspacing=0 cellpadding=3 width="95%">
  <form id=mtask_add name=mtask_add action=mtask_indb.asp?action=change method=post onSubmit='return checkinf();'>
    <tr>
      <th class=th height=25>��Ŀ����
        </td>
      <th class=th>��Ŀ����
        </td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>����ͬ��Ϣ��</b></td>
      <td class=ctd>&nbsp;</td>
    </tr>
    <tr>
      <td class=rtd width="20%">������</td>
      <td class=ltd><input type=text name=ddh size=30 value=<%=Rs("ddh")%>></td>
    </tr>
    <tr>
      <td class=rtd>��ˮ��</td>
      <td class=ltd><input type=text name=lsh size=30 value=<%=Rs("lsh")%>></td>
    </tr>
    <tr>
      <td class=rtd>ģ��</td>
      <td class=ltd><input type=text name=mh size=30 value=<%=Rs("mh")%>></td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td class=ltd><input type=text name=dmmc size=30 value="<%=Rs("dmmc")%>">
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
      <td class=rtd>�豸����</td>
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
      <td class=rtd>�������ͺ�</td>
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
      <td class=rtd>ģ�߲���</td>
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
      <td class=rtd>ˮ��ͷ����</td>
      <td class=ltd><input type=text name=sjtsl size=30 value=<%=Rs("sjtsl")%>></td>
    </tr>
    <tr>
      <td class=rtd>����ͷ����</td>
      <td class=ltd><input type=text name=qjtsl size=30 value=<%=Rs("qjtsl")%>></td>
    </tr>
    <tr>
      <td class=rtd>ǣ���ٶ�</td>
      <td class=ltd><input type=text name=qysd size=10 value=<%=Rs("qysd")%>>
        ��/��(m/min)</td>
    </tr>
    <tr>
      <td class=rtd>��������</td>
      <td class=ltd><select name=jcfx onchange=calmjfz();>
          <option value="/" <%If Rs("jcfx")="/" Then%> selected <%End If%>>&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <option value="��������" <%If Rs("jcfx")="��������" Then%> selected <%End If%>>��������</option>
          <option value="�ͻ�����" <%If Rs("jcfx")="�ͻ�����" Then%> selected <%End If%>>�ͻ�����</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>ǻ��</td>
      <td class=ltd><input type=text size=4 name=qs onchange=calmjfz(); value=<%=rs("qs")%>>ǻ</td>
    </tr>
    <tr>
      <td class=rtd>����Ȱ�</td>
      <td class=ltd><select name="pjrb">
          <option value=true<%If Rs("pjrb") Then%> selected<%End If%>>��</option>
          <option value=false<%If not(Rs("pjrb")) Then%> selected<%End If%>>��</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>���Ȱ���Ϣ</td>
      <td class=ltd> ����:
        <select name="jrbxs">
          <option value="����"<%If Rs("jrbxs")="����" Then%> selected<%End If%>>����</option>
          <option value="����"<%If Rs("jrbxs")="����" Then%> selected<%End If%>>����</option>
        </select>
        &nbsp;
        ����:
        <select name="jrbcl">
          <option value="����" <%If Rs("jrbcl")="����" Then%> selected <%End If%>>����</option>
          <option value="��ĸ"<%If Rs("jrbcl")="��ĸ" Then%> selected <%End If%>>��ĸ</option>
        </select>
        &nbsp;
        ����˵��:
        <input type=text name=jrbxx size=40 value=<%=Rs("jrbxx")%>></td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>����ֵ��Ϣ��</b></td>
      <td class=ctd></td>
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
      <td class=ltd><input name=fzxs type=text onchange="calmjfz()" size=5 value=<%=Rs("fzxs")%>></td>
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
    <tr>
      <td class=rtd>ģ�߷�ֵ</td>
      <td class=ltd> ģ���ܷ�:<span id=span_mjzf style="font-weight:bold;"><%=Rs("mjzf")%></span>��&nbsp;&nbsp;&nbsp;&nbsp; <span id=span_gjzf style="font-weight:bold;"> ����: <%=Rs("gjzf")%>��</span> <br>
        BOM�ܷ�:<span id=span_bomzf style="font-weight:bold;"><%=Rs("bomzf")%></span>��<br>
        �����ֲ��ܷ�:<span id=span_tsdzf style="font-weight:bold;"><%=Rs("tsdzf")%></span>��<br>
        �����ܷ�:<span id=span_tszf style="font-weight:bold;"><%=Rs("tszf")%></span>��<br>
        ������Ϣ�����ܷ�:<span id=span_tsxxzlzf style="font-weight:bold;"><%=Rs("tsxxzlzf")%></span>��<br></td>
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
      <td class=rtd rowspan="2">��ֵ����</td>
      <td class=ltd>ģͷ����:
        <input type=text name=mtbl size=4 onchange=blchange(); value=<%=Rs("mtbl")%>>
        %&nbsp;&nbsp;&nbsp;���ͱ���:
        <input type=text name=dxbl size=4 disabled value=<%=100-rs("mtbl")%>>
        %</td>
    </tr>
    <tr>
        <td class=ltd>ģͷ�ṹ:
        <input type=text name=mtjgbl size=4 value=<%=Rs("mtjgbl")%>>
        %&nbsp;&nbsp;&nbsp;���ͽṹ:
        <input type=text name=dxjgbl size=4 value=<%=Rs("dxjgbl")%>> %</td>
    </tr>
    <tr bgcolor="#DDDDDD">
      <td class=rtd height=25><b>��ģ����Ϣ��</b></td>
      <td class=ctd>&nbsp;</td>
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
      <td class=ltd><select name=rwlr>
          <option value="���" <%If Rs("rwlr")="���" Then%> selected <%End If%>>���</option>
          <option value="����" <%If Rs("rwlr")="����" Then%> selected <%End If%>>����</option>
          <option value="����" <%If Rs("rwlr")="����" Then%> selected <%End If%>>����</option>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>���ڵ���</td>
      <td class=ltd><select name="cnts" style="width:51px;" onchange='ExcTslb(this);'>
          <option value=true<%if rs("cnts") then%> selected<%end if%>>��</option>
          <option value=false<%if not(rs("cnts")) then%> selected<%end if%>>��</option>
        </select></td>
    </tr>
    <tr id=trbeit>
      <td class=rtd>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</td>
      <td class=ltd><select name="beit" style="width:51px;">
          <option value=true <%if rs("beit") then%> selected<%end if%>>��</option>
          <option value=false <%if not(rs("beit")) then%> selected<%end if%>>��</option>
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
      <td class=ltd><input type=text name=xcbh size=14 value=<%=Rs("xcbh")%>></td>
    </tr>
    <tr id=trdxqg>
      <td class=rtd>�����и�</td>
      <td class=ltd><select name="dxqg">
          <option value="���ϸ�"<%If Rs("dxqg")="���ϸ�" Then%> selected<%End If%>>���ϸ�</option>
          <option value="����ϸ�"<%If Rs("dxqg")="����ϸ�" Then%> selected<%End If%>>����ϸ�</option>
          <option value="����ϸ�"<%If Rs("dxqg")="����ϸ�" Then%> selected<%End If%>>����ϸ�</option>
          <option value="����һ���и�"<%if rs("dxqg")="����һ���и�" then%> selected<%end if%>>����һ���и�
        </select></td>
    </tr>
    <tr id=trmtjg>
      <td class=rtd>ģͷ�ṹ</td>
      <td class=ltd><input type=text name=mtjg size=30 value=<%=Rs("mtjg")%>>
        </td>
    </tr>
    <tr id=trdxjg>
      <td class=rtd>���ͽṹ</td>
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
      <td class=rtd>ˮ��ṹ</td>
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
      <td class=rtd>ģͷ���ӳߴ�</td>
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
      <td class=rtd>�ȵ�ż���</td>
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
      <td class=rtd height=25><b>��������Ϣ��</b></td>
      <td class=ctd>&nbsp;</td>
    </tr>
    <tr>
      <td class=rtd>��ע</td>
      <td class=ltd><textarea name="bz" cols="75" rows="7"><%=Rs("bz")%></textarea></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ���ʼʱ��</td>
      <td class=ltd><%
		dim Tmpshsj
		If IsNull(Rs("jhkssj")) Then Tmpshsj=now else Tmpshsj=Rs("jhkssj")
			%>
        <script language=javascript>
 		var myDate=new dateSelector(<%=year(Tmpshsj)&","&month(Tmpshsj)&","&day(Tmpshsj)%>);
  		myDate.year;
 		myDate.inputName='jhkssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ��ṹ����ʱ��</td>
      <td class=ltd><%
		dim Tmpjssj
		If IsNull(Rs("jhjgsj")) Then Tmpjssj=Rs("jhjssj") else Tmpjssj=Rs("jhjgsj")
			%>
        <script language=javascript>
 		var myDate=new dateSelector(<%=year(Tmpjssj)&","&month(Tmpjssj)&","&day(Tmpjssj)%>);
  		myDate.year;
 		myDate.inputName='jgjssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <tr>
      <td class=rtd>�ƻ�ȫ�׽���ʱ��</td>
      <td class=ltd><%
		If IsNull(Rs("jhjssj")) Then Tmpjssj=now else Tmpjssj=Rs("jhjssj")
			%>
        <script language=javascript>
 		var myDate=new dateSelector(<%=year(Tmpjssj)&","&month(Tmpjssj)&","&day(Tmpjssj)%>);
  		myDate.year;
 		myDate.inputName='jhjssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script></td>
    </tr>
    <% '13:40 2007-1-6-������
'������ǰֻ��һ������ֻ��һ���鳤�İ汾
If Rs("zz") <> "" Then %>
    <tr>
      <td class=rtd>�鳤</td>
      <td class=ltd><select name="zz">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%If Rs("zz")=c_allzz(i) Then%> selected<%End If%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <%else%>
    <tr>
      <td class=rtd>�ṹ�鳤</td>
      <td class=ltd><select name="jgzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%if rs("jgzz")=c_allzz(i) then%> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=rtd>����鳤</td>
      <td class=ltd><select name="sjzz" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>'<%if rs("sjzz")=c_allzz(i) then%> selected<%end if%>><%=c_allzz(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <%End If%>
    <tr>
      <td class=rtd>��������</td>
      <td class=ltd><select name="jsdb" style="width:80px;">
          <option></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'<%If Rs("jsdb")=c_allzy(i) Then%> selected<%End If%>><%=c_allzy(i)%></option>
          <%next%>
        </select></td>
    </tr>
    <tr>
      <td class=ctd colspan=2><input type=submit value=" �� ȷ �� �� "></td>
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
<%Call TbTopic("������ˮ�� <font style=color:#0000FF>" & Rs("lsh") & "</font> �ı�ע")%>
<%If ChkAble(4) Then Response.Write "<a href=mtask_zzchange.asp?s_lsh="&Rs("lsh")&">�鳤Ȩ��</a><br>"%>
<table class=xtable cellspacing=0 cellpadding=10>
  <form id=Bz_change name=Bz_change action=mtask_indb.asp?action=BzChan method=post onSubmit='return true;'>
	<tr>
		<td class=ctd >������</td>
		<td class=ctd ><%=Rs("ddh")%></td>
		<td class=ctd >�ͻ���</td>
		<td class=ctd ><%=Rs("dwmc")%></td>
	</tr>
    <tr>
      <td class=ctd colspan="4"><textarea name="bz" cols="75" rows="10"><%=Rs("bz")%></textarea></td>
    </tr>
    <tr>
      <td class=ctd colspan="4"><input type=submit value=" �� ȷ �� �� "></td>
    </tr>
    <input type="hidden" name=id value=<%=Rs("id")%>>
  </form>
</table>
<%
End Function		'BzChan()

Function mtask_js(CkdmOv,TslbOv,DmdeOV)
'����ΪJS����%>
<script language="javascript">
//��ʼ��������������ĳ������
<%If Rs("cnts") Then%>
	document.all.trtslb.style.display="";
	document.all.trbeit.style.display="none";
<%else%>
	document.all.trtslb.style.display="none";
	document.all.trbeit.style.display="";
<%End If%>

//���ڵ��������˵���������
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

//�Բο�ģ�߳�ʼ��
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
	}
 	var   o   =   document.getElementById("tslb");
	var   i;
	for(i=0;i<o.length;i++)
	{
 		if(o.options[i].value=="<%=TslbOv%>")
 			o.options[i].selected=true;
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
	if(isNaN(parseFloat(document.all.mtjgbl.value))) document.all.mtjgbl.value=mtjgbl*100;
	if(isNaN(parseFloat(document.all.dxjgbl.value))) document.all.dxjgbl.value=dxjgbl*100;

	//���������ֵ
	var i_gjfs=0;
	var i_qhgj=0;

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

//ģͷ��ֵ����
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
			break;
		default:
			document.all.trdxqg.style.display="";
			document.all.trdxjg.style.display="";
			document.all.trsxjg.style.display="";
			break;
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
calmjfz();
chkmjxx(document.all.mjxx);
</script>
<%
End Function	'mtask_js()
%>
