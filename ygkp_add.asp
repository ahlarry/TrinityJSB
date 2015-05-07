<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="员工考评 → 添加员工考评"
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

Function kpHead() '考评页面页首
%>
<%
	If Chkable(2) Then		'技术厂长
	%>
技术厂长 → <a href="?kp=cztozrkp">主任考评</a><br>
<%
	End If

	If Chkable(3) Then		'主任
	%>
主&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;任 → <a href="?kp=zrtozykp">设计技术员考评</a> | <a href="?kp=zrtofwkp">服务技术员考评</a> <br>
<%
	End If

	If Chkable(11) Then	'品管部
	%>
品&nbsp;管&nbsp;&nbsp;部 → <a href="?kp=PgbToTsykp">调试技术员考评</a> | <a href="?kp=pgbtozykp">设计、服务技术员考评</a><br>
<%
	End If

	If Chkable(12) Then	'管理部
	%>
管&nbsp;理&nbsp;&nbsp;部 → <a href="?kp=glbtozrkp">技术部主任考评</a>
<%
	End If
End Function

Function CzToZrKpAdd()	 '厂长to主任考评
	Call ChkPageAble(2)
%>
<%Call TbTopic("技术厂长 → 主任考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>被考评人员</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_alljl)%>
          <option value='<%=c_alljl(i)%>'><%=c_alljl(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
					strSql="Select * from [c_kplr] where kp_kind=1 and kp_kpr=1"	'kpr=1为技术厂长
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
      <td class=th>备注</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 确定 "></td>
    </Tr>
    <input type="hidden" name="action" value="cztozrkp">
  </form>
</Table>
<%
End Function

Function ZrToZyKpAdd()		'主任 to 组员考评
	Call ChkPageAble(3)
%>
<%Call TbTopic("主任 → 设计技术员考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form name="zrtozy" action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>被考评人员</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><Select name="kpinfo" onChange="changexm();">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=5 and kp_kpr=2 or kp_kind=5 and kp_kpr=3 or kp_kind=5 and kp_kpr=106 order by kp_topic desc"		'kp_kind=5代表组员,kp_kpr=2代表主任
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
      <td class=th>考评分值</td>
      <td class=ltd><label>
          <input name="kpfz" type="text" id="kpfz" value="0" readonly="true" onKeyPress="javascript:validationNumber(this, 'float', '', '');" />
        </label></td>
    </tr>
    <tr>
      <td class=th>备注</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 确定 "></td>
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
			if(TmpArr[1]=="其他")
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


Function ZrToFwKpAdd()		'主任 to 服务技术员
	Call ChkPageAble(3)
%>
<%Call TbTopic("主任 → 服务技术员考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>被考评人员</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_allgy)%>
          <option value='<%=c_allgy(i)%>'><%=c_allgy(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><Select name="kpinfo" onChange="changexm();">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=8 and kp_kpr=2 or kp_kind=8 and kp_kpr=3 or kp_kind=8 and kp_kpr=106 order by kp_id"		'kp_kind=5代表组员,kp_kpr=2代表主任
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
      <td class=th>考评分值</td>
      <td class=ltd><label>
          <input name="kpfz" type="text" id="kpfz" value="0" readonly="true" onKeyPress="javascript:validationNumber(this, 'float', '', '');" />
        </label></td>
    </tr>
    <tr>
      <td class=th>备注</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 确定 "></td>
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

Function ZzToTsykpAdd()		'组长 to 调试员考评
	Call ChkPageAble(4)
	If Session("usergroup")<>5 Then Exit Function
%>
<%Call TbTopic("组长 → 调试员考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="70%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <input type="hidden" name="kpzrr" value="">
    <Tr>
      <td class=th width=100>被考评人员</td>
      <td class=ltd width="*"><span id="kpzrr_dis" name="kpzrr_dis">请选择被考评人</span></td>
    </tr>
    <tr>
      <td class=th>被考评人列表</td>
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
				document.all.kpzrr_dis.innerHTML="请选择被考评人";
			else
				document.all.kpzrr_dis.innerHTML=strtemp;
			document.all.kpzrr.value=strtemp;
		}

	</script>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=4 and kp_kpr=3"		'3代表组长
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
      <td class=th>模具流水号</td>
      <td class=ltd><input type=text name="hglsh" size=15>
        </span>(除样品合格外，此项均不填)</td>
    </Tr>
    <tr>
      <td class=th>备注</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 确定 "></td>
    </Tr>
    <input type="hidden" name="action" value="zztotsykp">
  </form>
</Table>
<%
End Function


Function ZzToZykpAdd()		'组员考评
	Call ChkPageAble(4)
%>
<%Call TbTopic("组长 → 组员考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>被考评人员</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(eval("c_xz" & Session("usergroup")))%>
          <option value='<%=eval("c_xz" & Session("usergroup"))(i)%>'><%=eval("c_xz" & Session("usergroup"))(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=5 and kp_kpr=3"		'3代表组长
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
      <td class=th>备注</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 确定 "></td>
    </Tr>
    <input type="hidden" name="action" value="zztozykp">
  </form>
</Table>
<%
End Function

Function PgbToTsykpAdd()	'品管部 to  调试技术员考评
	Call ChkPageAble(11)
%>
<%Call TbTopic("品管部 →  调试技术员考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="70%"  align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkpgbkp(this);">
    <input type="hidden" name="kpzrr" value="">
    <Tr>
      <td class=th width=100>被考评人</td>
      <td class=ltd width="*"><span id="kpzrr_dis" name="kpzrr_dis">请选择被考评人</span></td>
    </tr>
    <tr>
      <td class=th>调试人列表</td>
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
				document.all.kpzrr_dis.innerHTML="请选择被考评人";
			else
			{
				if(strtempSH=="")
					document.all.kpzrr_dis.innerHTML=strtemp;
				else
					document.all.kpzrr_dis.innerHTML=strtemp + "||" + strtempSH + "(审核)";
			}
			document.all.kpzrr.value=strtemp;
		}

	</script>
    <Tr>
      <td class=th width=100>审核人</td>
      <td class=ltd><select name="kpsh" onChange="changetsy();">
          <Option value=""></Option>
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </select>
        &nbsp;&nbsp;&nbsp;(可选) </td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=4 and kp_kpr=4 or kp_kind=4 and kp_kpr=106"		'4代表品管部
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
      <td class=th width=100>零件件数</td>
      <td class=ltd><input type="text" name="ljjs" size=8 value="1.0"></td>
    </tr>
    <Tr>
      <td class=th width=100>零件系数</td>
      <td class=ltd><input type="text" name="ljxs" size=8 value="1.0"></td>
    </tr>
    <tr>
      <td class=th>备注</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 确定 "></td>
    </Tr>
    <input type="hidden" name="action" value="PgbToTsykp">
  </form>
</Table>
<%
End Function

Function PgbToZyKpAdd()		'品管部 to 组员考评
	Call ChkPageAble(11)
%>
<%Call TbTopic("品管部 → 设计技术员考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%" align="center">
  <form name="Pg2Zy" action="ygkp_indb.asp" method="post" onSubmit="return chkpgbkp(this);">
    <Tr>
      <td class=th width=100>模具流水号</td>
      <td colspan="2" class=ltd><input type="text" name="kplsh" size=15></td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td colspan="2" class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=5 and kp_kpr=4 or kp_kind=5 and kp_kpr=106"		'3代表组长
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
      <td width="100" rowspan="3" class=th>被考评人员</td>
      <td class=ltd>设计:
        <input name="kpsj" type="text" id="kpsj" size="25" readonly=true />
        <a href="javascript:zrrminus('kpsj');"><img src="images/undo.png" alt="重置" border=0 align="middle"></a>
        <a href="javascript:zrrplus('kpsj');"><img src="images/plus.png" alt="添加" border=0 align="middle"></a>
      </td>
      <td rowspan="3" class=ltd><select name="hxzrr" size="6" multiple="multiple" id="hxzrr">
          <%for i = 0 to ubound(c_allzy)%>
          <option value='<%=c_allzy(i)%>'><%=c_allzy(i)%></option>
          <%next%>
        </select>
        </label></td>
    </Tr>
    <Tr>
      <td class=ltd>审核:
        <input name="kpsh" type="text" id="kpsh" size="25" readonly=true />
        <a href="javascript:zrrminus('kpsh');"><img src="images/undo.png" alt="重置" border=0 align="middle"></a>
        <a href="javascript:zrrplus('kpsh');"><img src="images/plus.png" alt="添加" border=0 align="middle"></a>
      </td>
    </Tr>
    <Tr>
      <td class=ltd>
        鼠标拖动连续选择或按住Ctrl鼠标点击间隔多选
      </td>
    </Tr>
    <Tr>
      <td class=th width=100>零件件数</td>
      <td colspan="2" class=ltd><input type="text" name="ljjs" size=8 value="1.0"></td>
    </tr>
    <Tr>
      <td class=th width=100>零件系数</td>
      <td colspan="2" class=ltd><input type="text" name="ljxs" size=8 value="1.0"></td>
    </tr>
    <tr>
      <td class=th>备注</td>
      <td colspan="2" class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td colspan="2" class=ltd><input type="submit" value=" 确定 "></td>
    </Tr>
    <input type="hidden" name="action" value="pgbtozykp">
  </form>
</Table>
<%
End Function

Function GlbKpAdd()		'管理部考评
	Call ChkPageAble(12)
%>
<%Call TbTopic("管理部 → 技术部主任考评")%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%" align="center">
  <form action="ygkp_indb.asp" method="post" onSubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>被考评人员</td>
      <td class=ltd><Select name="kpzrr">
          <option value=""></option>
          <%for i = 0 to ubound(c_alljl)%>
          <option value='<%=c_alljl(i)%>'><%=c_alljl(i)%></option>
          <%next%>
        </Select></td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><Select name="kpinfo">
          <option value=""></option>
          <%
						strSql="Select * from c_kplr where kp_kind=1 and kp_kpr=5"		'3代表组长
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
      <td class=th>备注</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 确定 "></td>
    </Tr>
    <input type="hidden" name="action" value="glbtozrkp">
  </form>
</Table>
<%
End Function
%>
<script language="javascript">
function getNoRepeat(arg)		//过滤以,分隔的字符串中重复项
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