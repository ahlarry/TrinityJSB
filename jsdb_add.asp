<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->

<%
Call ChkPageAble(3)
CurPage="设计任务 → 添加技术代表任务"
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
<table  class="xtable" cellspacing="0" cellpadding="0" width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd height=300><%Call ftaskCadd()%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</table>
<%
End Sub

function ftaskCadd()
	Dim s_hth,s_khmc,s_rwnr,s_fz,s_sj,s_zz
	s_hth="" : s_khmc="" : s_rwnr="" : s_fz=0 : s_sj="" : s_zz=""
	If Trim(request("shth"))<>"" Then s_hth=Trim(request("shth"))
	if s_hth<>"" Then
		strSql="select * from [jsdb] where hth='"&s_hth&"'"
		Set Rs=xjweb.Exec(strSql,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("合同号 " & s_hth & " 任务书不存在! 请重新输入!","jsdb_add.asp")
		Else
			If Not isnull(rs("shjssj")) then
				Call JsAlert("本任务已经完成,无法更改任务书","jsdb_add.asp")
			else
				s_khmc=Rs("khmc")
				s_rwnr=Rs("rwnr")
				s_fz=Rs("jcf")
				s_sj=Rs("jhjssj")
				s_zz=Rs("zz")
			end if
		end if
	End  if
%>
<%Call TbTopic("添加修改技术代表任务")%>
<form id="frm_ftask" name="ftask_add" action="jsdb_indb.asp?action=add" method="post" onsubmit='return checkinf();'>
 <table  id="table1" class="xtable" cellspacing="0" cellpadding="3" width="80%">
		<tr>
		<td class="rtd">合同号</td>
        <td class="ltd"><input type="text" name="hth" size="23" value=<%=s_hth%>><button id="chg" onClick='location.href("<%=request.servervariables("script_name")%>?shth="+this.form.hth.value);'>修改</button></td>
    </tr>
    <tr>
      <td class="rtd">客户名称</td>
      <td class="ltd"><input type="text" name="khmc" size="30" value=<%=s_khmc%>></td>
    </tr>
	<tr>
      <td class="rtd">任务内容</td>
      <td class="ltd"><input name="db1" type=checkbox id="db1" onclick="checkxz();" value="消化合同审核任务书" <%If  InStr(s_rwnr,"消化合同审核任务书")>0 Then%> checked <%End If%> >
        消化合同，审核任务书
        <input type=checkbox id="db2" name="db2" value="配置接口件" onclick="checkxz();" <%If  InStr(s_rwnr,"配置接口件")>0 Then%> checked <%End If%>>
        配置接口件
        <input type=checkbox id="db3" name="db3" value="带共挤" onclick="checkxz();" <%If  InStr(s_rwnr,"带共挤")>0 Then%> checked <%End If%>>
        带共挤
        <input type=checkbox id="db4" name="db4" value="配置非标水箱" onclick="checkxz();" <%If  InStr(s_rwnr,"配置非标水箱")>0 Then%> checked <%End If%>>
        配置非标水箱</td>
    </tr>
    <tr>
      <td class="rtd">分值</td>
      <td class="ltd"><input type="text" id="jcf" name="jcf" size="8"  value=<%=s_fz%> />
        </td>
    </tr>
    <tr>
      <td class=rtd>计划结束时间</td>
      <td colspan="2" class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jhjssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script><font color="#ff0000"><strong>修改任务书时,此时间需重新选择</strong></font></td>
    </tr>
    <tr>
      <td class="rtd">责任组长</td>
      <td class="ltd"><select name="sjr">
          <option></option>
          <%for i = 0 to ubound(c_allzz)%>
          <option value='<%=c_allzz(i)%>' <%if s_zz=c_allzz(i) Then%>selected<%End If%>><%=c_allzz(i)%></option>
          <%next%>
        </select>
      </td>
    </tr>
</table>
<table class="xtable" cellspacing="0" cellpadding="3" width="80%">
<tr>
      <td class="ctd" colspan="2"><input type="hidden" name="rwnr" ><input type="submit" value=" · 确 定 · " /></td>
    </tr>
</table>
</form>
<%
end function	
%>
<script language="javascript">
		function checkxz()
		{
			var rwnr="";
			for(i=1;i<5;i++) {
				var chkobj=eval("document.all.db" + i);
				if(chkobj.checked){
					rwnr=rwnr + "," + chkobj.value;
				}
				}
				document.all.rwnr.value=rwnr.substring(1);
		}
		function checkinf()
	{
		if (document.all.jcf.value==0){alert("任务分值不能为0！\n");document.all.jcf.focus();return false;}
	}
</script>