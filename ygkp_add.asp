<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="Ա������ �� ���Ա������"
strPage="ygkp"
Call FileInc(0, "js/ygkp.js")
xjweb.header()
Call TopTable()

Dim action
action=Request("kp")
Call Main()

Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<body>
<Table class=xtable cellspacing=0 cellpadding=4 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ltd height=25><%Call kpHead()%></Td>
  </Tr>
  <Tr>
    <Td class=ctd height=300><%
				Select Case action
					Case "cztozrkp"
						Call CzToZrKpAdd()
					Case "zrtozykp"
						Call ZrToZyKpAdd()
					Case "zrtofwkp"
						Call ZrToFwkpAdd()
					Case "zztotsykp"
						Call ZzToTsyKpAdd()
					Case "zztozykp"
						Call ZzToZyKpAdd()
					Case "PgbToTsykp"
						Call PgbToTsykpAdd()
					Case "pgbtozykp"
						Call PgbToZyKpAdd()
					Case "glbtozrkp"
						Call GlbKpAdd()
					Case Else
				End Select
			%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function kpHead() '����ҳ��ҳ��
%>
<%
	If Chkable(2) Then		'��������
	%>
�������� �� <a href="?kp=cztozrkp">���ο���</a><br>
<%
	End If

	If Chkable(3) Then		'����
	%>
��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� �� <a href="?kp=zrtozykp">��Ƽ���Ա����</a> | <a href="?kp=zrtofwkp">������Ա����</a> <br>
<%
	End If

	If Chkable(11) Then	'Ʒ�ܲ�
	%>
Ʒ&nbsp;��&nbsp;&nbsp;�� �� <a href="?kp=PgbToTsykp">���Լ���Ա����</a> | <a href="?kp=pgbtozykp">��ơ�������Ա����</a><br>
<%
	End If

	If Chkable(12) Then	'����
	%>
��&nbsp;��&nbsp;&nbsp;�� �� <a href="?kp=glbtozrkp">���������ο���</a>
<%
	End If
End Function

Function CzToZrKpAdd()	 '����to���ο���
	Call ChkPageAble(2)
%>
<%Call TbTopic("�������� �� ���ο���")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>��������Ա</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_alljl)%>
          <option value='<%=c_alljl(i)%>'><%=c_alljl(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
					strSql="Select * from [c_kplr] where kp_kind=1 and kp_kpr=1"	'kpr=1Ϊ��������
					Set Rs=xjweb.Exec(strSql, 1)
					Do While Not Rs.Eof
						%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_topic")%></Option>
          <%
						Rs.MoveNext
					Loop
				%>
        </Select></td>
    </tr>
    <tr>
      <td class=th>��ע</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="cztozrkp">
  </form>
</Table>
<%
End Function

Function ZrToZyKpAdd()		'���� to ��Ա����
	Call ChkPageAble(3)
%>
<%Call TbTopic("���� �� ��Ƽ���Ա����")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form name="zrtozy" action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>��������Ա</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td class=ltd><Select name="kpinfo" onChange="changexm();">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=5 and kp_kpr=2 or kp_kind=5 and kp_kpr=3 or kp_kind=5 and kp_kpr=106 order by kp_topic desc"		'kp_kind=5������Ա,kp_kpr=2��������
						Set Rs=xjweb.Exec(strSql, 1)
						Do While Not Rs.Eof
							%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_item")%></Option>
          <%
							Rs.MoveNext
						Loop
					%>
        </Select></td>
    </tr>
    <tr id=trkpfz>
      <td class=th>������ֵ</td>
      <td class=ltd><label>
          <input name="kpfz" type="text" id="kpfz" value="0" readonly="true" onKeyPress="javascript:validationNumber(this, 'float', '', '');" />
        </label></td>
    </tr>
    <tr>
      <td class=th>��ע</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="zrtozykp">
  </form>
</Table>
<script language="javascript">
		function changexm()
		{
			var TmpArr = new Array();
			var str1 = document.all.kpinfo.value;
			TmpArr = str1.split("||")
			document.all.kpfz.value=TmpArr[2];
			if(TmpArr[1]=="����")
				{
				document.getElementById("kpfz").readOnly=false;
				}
			else
				{
				document.getElementById("kpfz").readOnly=true;
				}
		}
	</script>
<%
End Function


Function ZrToFwKpAdd()		'���� to ������Ա
	Call ChkPageAble(3)
%>
<%Call TbTopic("���� �� ������Ա����")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>��������Ա</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_allgy)%>
          <option value='<%=c_allgy(i)%>'><%=c_allgy(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td class=ltd><Select name="kpinfo" onChange="changexm();">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=8 and kp_kpr=2 or kp_kind=8 and kp_kpr=3 or kp_kind=8 and kp_kpr=106 order by kp_id"		'kp_kind=5������Ա,kp_kpr=2��������
						Set Rs=xjweb.Exec(strSql, 1)
						Do While Not Rs.Eof
							%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_item")%></Option>
          <%
							Rs.MoveNext
						Loop
					%>
        </Select></td>
    </tr>
    <tr id=trkpfz>
      <td class=th>������ֵ</td>
      <td class=ltd><label>
          <input name="kpfz" type="text" id="kpfz" value="0" readonly="true" onKeyPress="javascript:validationNumber(this, 'float', '', '');" />
        </label></td>
    </tr>
    <tr>
      <td class=th>��ע</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="zrtogykp">
  </form>
</Table>
<script language="javascript">
		function changexm()
		{
			var TmpArr = new Array();
			var str1 = document.all.kpinfo.value;
			TmpArr = str1.split("||")
			document.all.kpfz.value=TmpArr[2];
			if(TmpArr[2]==0)
				{
				document.getElementById("kpfz").readOnly=false;
				}
			else
				{
				document.getElementById("kpfz").readOnly=true;
				}
		}
	</script>
<%
End Function

Function ZzToTsykpAdd()		'�鳤 to ����Ա����
	Call ChkPageAble(4)
	If Session("usergroup")<>5 Then Exit Function
%>
<%Call TbTopic("�鳤 �� ����Ա����")%>
<table class=xtable cellspacing=0 cellpadding=3 width="70%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <input type="hidden" name="kpzrr" value="">
    <Tr>
      <td class=th width=100>��������Ա</td>
      <td class=ltd width="*"><span id="kpzrr_dis" name="kpzrr_dis">��ѡ�񱻿�����</span></td>
    </tr>
    <tr>
      <td class=th>���������б�</td>
      <td class=ltd><table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <%
				dim j
				j=1
				for i=0 to ubound(c_alltsy)
					if j>10 then response.write("</tr><tr>") : j=1
			%>
            <td><input type=checkbox id=user<%=i%> name=user<%=i%> value=<%=c_alltsy(i)%> class=radio onClick="changetsy();">
              <label for=user<%=i%>><%=c_alltsy(i)%></label></td>
            <%
					j=j+1
				next
			%>
          </tr>
        </table></td>
    </tr>
    <script language="javascript">
		function changetsy()
		{
			var ii=0;
			var strtemp="";
			for(ii=0; ii<=<%=ubound(c_alltsy)%>; ii++)
			{
				if(eval("document.all.user" + ii +".checked==true"))
					//alert(eval("document.all.user" + ii + ".value"));
				{
					if(strtemp!="")
						strtemp=strtemp + "|" + eval("document.all.user" + ii + ".value");
					else
						strtemp=eval("document.all.user" + ii + ".value");
				}
			}

			if(strtemp=="")
				document.all.kpzrr_dis.innerHTML="��ѡ�񱻿�����";
			else
				document.all.kpzrr_dis.innerHTML=strtemp;
			document.all.kpzrr.value=strtemp;
		}

	</script>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=4 and kp_kpr=3"		'3�����鳤
						Set Rs=xjweb.Exec(strSql, 1)
						Do While Not Rs.Eof
							%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_item")%></Option>
          <%
							Rs.MoveNext
						Loop
					%>
        </Select></td>
    </tr>
    <Tr>
      <span id=tshglsh style=display:black;>
      <td class=th>ģ����ˮ��</td>
      <td class=ltd><input type=text name="hglsh" size=15>
        </span>(����Ʒ�ϸ��⣬���������)</td>
    </Tr>
    <tr>
      <td class=th>��ע</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="zztotsykp">
  </form>
</Table>
<%
End Function


Function ZzToZykpAdd()		'��Ա����
	Call ChkPageAble(4)
%>
<%Call TbTopic("�鳤 �� ��Ա����")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>��������Ա</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(eval("c_xz" & Session("usergroup")))%>
          <option value='<%=eval("c_xz" & Session("usergroup"))(i)%>'><%=eval("c_xz" & Session("usergroup"))(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=5 and kp_kpr=3"		'3�����鳤
						Set Rs=xjweb.Exec(strSql, 1)
						Do While Not Rs.Eof
							%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_item")%></Option>
          <%
							Rs.MoveNext
						Loop
					%>
        </Select></td>
    </tr>
    <tr>
      <td class=th>��ע</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="zztozykp">
  </form>
</Table>
<%
End Function

Function PgbToTsykpAdd()	'Ʒ�ܲ� to  ���Լ���Ա����
	Call ChkPageAble(11)
%>
<%Call TbTopic("Ʒ�ܲ� ��  ���Լ���Ա����")%>
<table class=xtable cellspacing=0 cellpadding=3 width="70%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkpgbkp(this);">
    <input type="hidden" name="kpzrr" value="">
    <Tr>
      <td class=th width=100>��������</td>
      <td class=ltd width="*"><span id="kpzrr_dis" name="kpzrr_dis">��ѡ�񱻿�����</span></td>
    </tr>
    <tr>
      <td class=th>�������б�</td>
      <td class=ltd><table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <%
				dim j
				j=1
				for i=0 to ubound(c_alltsy)
					if j>10 then response.write("</tr><tr>") : j=1
			%>
            <td><input type=checkbox id=user<%=i%> name=user<%=i%> value=<%=c_alltsy(i)%> class=radio onClick="changetsy();">
              <label for=user<%=i%>><%=c_alltsy(i)%></label></td>
            <%
					j=j+1
				next
			%>
          </tr>
        </table></td>
    </tr>
    <script language="javascript">
		function changetsy()
		{
			var ii=0;
			var strtemp="";
			var strtempSH=document.all.kpsh.value;
			for(ii=0; ii<=<%=ubound(c_alltsy)%>; ii++)
			{
				if(eval("document.all.user" + ii +".checked==true"))
					//alert(eval("document.all.user" + ii + ".value"));
				{
					if(strtemp!="")
						strtemp=strtemp + "|" + eval("document.all.user" + ii + ".value");
					else
						strtemp=eval("document.all.user" + ii + ".value");
				}
			}
			if(strtemp=="")
				document.all.kpzrr_dis.innerHTML="��ѡ�񱻿�����";
			else
			{
				if(strtempSH=="")
					document.all.kpzrr_dis.innerHTML=strtemp;
				else
					document.all.kpzrr_dis.innerHTML=strtemp + "||" + strtempSH + "(���)";
			}
			document.all.kpzrr.value=strtemp;
		}

	</script>
    <Tr>
      <td class=th width=100>�����</td>
      <td class=ltd><select name="kpsh" onChange="changetsy();">
          <Option value=""></Option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </select>
        &nbsp;&nbsp;&nbsp;(��ѡ) </td>
    </Tr>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=4 and kp_kpr=4 or kp_kind=4 and kp_kpr=106"		'4����Ʒ�ܲ�
						Set Rs=xjweb.Exec(strSql, 1)
						Do While Not Rs.Eof
							%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_item")%></Option>
          <%
							Rs.MoveNext
						Loop
					%>
        </Select></td>
    </tr>
    <Tr>
      <td class=th width=100>�������</td>
      <td class=ltd><input type="text" name="ljjs" size=8 value="1.0"></td>
    </tr>
    <Tr>
      <td class=th width=100>���ϵ��</td>
      <td class=ltd><input type="text" name="ljxs" size=8 value="1.0"></td>
    </tr>
    <tr>
      <td class=th>��ע</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="PgbToTsykp">
  </form>
</Table>
<%
End Function

Function PgbToZyKpAdd()		'Ʒ�ܲ� to ��Ա����
	Call ChkPageAble(11)
%>
<%Call TbTopic("Ʒ�ܲ� �� ��Ƽ���Ա����")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%" align="center">
  <form name="Pg2Zy" action="ygkp_indb.asp" method="post" onSubmit="return chkpgbkp(this);">
    <Tr>
      <td class=th width=100>ģ����ˮ��</td>
      <td colspan="2" class=ltd><input type="text" name="kplsh" size=15></td>
    </Tr>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td colspan="2" class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=5 and kp_kpr=4 or kp_kind=5 and kp_kpr=106"		'3�����鳤
						Set Rs=xjweb.Exec(strSql, 1)
						Do While Not Rs.Eof
							%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_item")%></Option>
          <%
							Rs.MoveNext
						Loop
					%>
        </Select></td>
    </tr>
    <Tr>
      <td width="100" rowspan="3" class=th>��������Ա</td>
      <td class=ltd>���:
        <input name="kpsj" type="text" id="kpsj" size="25" readonly=true />
        <a href="javascript:zrrminus('kpsj');"><img src="images/undo.png" alt="����" border=0 align="middle"></a>
        <a href="javascript:zrrplus('kpsj');"><img src="images/plus.png" alt="���" border=0 align="middle"></a>
      </td>
      <td rowspan="3" class=ltd><select name="hxzrr" size="6" multiple="multiple" id="hxzrr">
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </select>
        </label></td>
    </Tr>
    <Tr>
      <td class=ltd>���:
        <input name="kpsh" type="text" id="kpsh" size="25" readonly=true />
        <a href="javascript:zrrminus('kpsh');"><img src="images/undo.png" alt="����" border=0 align="middle"></a>
        <a href="javascript:zrrplus('kpsh');"><img src="images/plus.png" alt="���" border=0 align="middle"></a>
      </td>
    </Tr>
    <Tr>
      <td class=ltd>
        ����϶�����ѡ���סCtrl����������ѡ
      </td>
    </Tr>
    <Tr>
      <td class=th width=100>�������</td>
      <td colspan="2" class=ltd><input type="text" name="ljjs" size=8 value="1.0"></td>
    </tr>
    <Tr>
      <td class=th width=100>���ϵ��</td>
      <td colspan="2" class=ltd><input type="text" name="ljxs" size=8 value="1.0"></td>
    </tr>
    <tr>
      <td class=th>��ע</td>
      <td colspan="2" class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td colspan="2" class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="pgbtozykp">
  </form>
</Table>
<%
End Function

Function GlbKpAdd()		'��������
	Call ChkPageAble(12)
%>
<%Call TbTopic("���� �� ���������ο���")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%" align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>��������Ա</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_alljl)%>
          <option value='<%=c_alljl(i)%>'><%=c_alljl(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>������Ŀ</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=1 and kp_kpr=5"		'3�����鳤
						Set Rs=xjweb.Exec(strSql, 1)
						Do While Not Rs.Eof
							%>
          <Option value="<%=Rs("kp_topic") & "||" & Rs("kp_item") & "||" & Rs("kp_uprice") & "||" & Rs("Kp_mul")%>"><%=Rs("kp_item")%></Option>
          <%
							Rs.MoveNext
						Loop
					%>
        </Select></td>
    </tr>
    <tr>
      <td class=th>��ע</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>����</td>
      <td class=ltd><input type="submit" value=" ȷ�� "></td>
    </Tr>
    <input type="hidden" name="action" value="glbtozrkp">
  </form>
</Table>
<%
End Function
%>
<script language="javascript">
function getNoRepeat(arg)		//������,�ָ����ַ������ظ���
{
	var strArr = arg.split(",");
	var str = ","
	var strt=""
	for(i = 0; i < strArr.length; i++)
	{
		if(str.indexOf("," + strArr[i] + ",") == -1)str += strArr[i] + ","
	}
	strt = str.substring(1,str.length - 1);
	return strt;
}

function zrrplus(arg1)
{
	var i;
   	for (i=0;i<document.Pg2Zy.hxzrr.length;i++){
    	if (document.Pg2Zy.hxzrr.options[i].selected == true){
    		if (eval("document.Pg2Zy." + arg1 + ".value==''"))
    			eval("document.Pg2Zy." + arg1 + ".value=document.all.hxzrr.options[i].value;");
    		else
        		eval("document.Pg2Zy." + arg1 + ".value=document.Pg2Zy." + arg1 + ".value+','+document.all.hxzrr.options[i].value;");
       }
    }
    eval("var Tmps = getNoRepeat(document.Pg2Zy." + arg1 + ".value);");
    eval("document.Pg2Zy." + arg1 + ".value=Tmps;");
}

function zrrminus(arg) {
	var i;
	eval("document.Pg2Zy." + arg + ".value='';");
   	for (i=0;i<document.Pg2Zy.hxzrr.length;i++){
    	document.Pg2Zy.hxzrr.options[i].selected = false;
    }
}
</script>