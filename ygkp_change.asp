<!--#include file="include/conn.asp"-->
<!--#include file="include/calendar.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="员工考评 → 更改员工考评"
strPage="ygkp"
Call FileInc(0, "js/ygkp.js")
xjweb.header()
Call TopTable()
	Dim strZrr, strkpitem, strgzz, strkpjs, strclsh, striPage, iid
	strZrr=Trim(Request("zrr"))
	strkpitem = trim(request("kpitem"))
	strkpjs = trim(request("kpjs"))
	strclsh = trim(request("clsh"))
	strgzz =request("gzz")
	striPage =request("ipage")
	iid=Request("id")
	If Not IsNumeric(iid) Then Call JsAlert("请从正确入口进入!","index.asp")
	iid=CLng(iid)
	Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd height=300><%
				Call KpChange()
			%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function KpChange()
	Dim Tmpkpsj
%>
<%Call TbTopic("更改考评信息")%>
<%
	strSql="Select * from [kp_jsb] where id="&iid&""
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then	Call JsAlert("此信息不存在!请核实!","ygkp_list.asp")
	If Rs("kp_kpr")<>Session("userName") and not(ChkAble("3")) Then Call JsAlert("请联系 "&Rs("kp_kpr")&" 进行更改!","")
	%>
<table class=xtable cellspacing=0 cellpadding=3 width="60%">
  <form action="ygkp_indb.asp" method="post" onsubmit="return chkygkp(this);">
    <Tr>
      <td class=th width=100>考评人员</td>
      <td class=ltd><%=Rs("kp_zrr")%></td>
    </Tr>
    <Tr>
      <td class=th width=100>考评项目</td>
      <td class=ltd><%=Rs("kp_item")%></td>
    </tr>
    <Tr>
      <td class=th width=100>原考评时间</td>
      <td class=ltd><%=xjDate(rs("kp_time"),1)%></td>
    </tr>
    <tr>
      <td class=rtd>考评时间</td>
      <td colspan="2" class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='kpsj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script></td>
    </tr>
    </tr>
    <Tr>
      <td class=th width=100>考评分</td>
      <td class=ltd><input type="text" name="kpfz" value="<%=Rs("kp_uprice")%>" onkeypress="javascript:validationNumber(this, 'float', 10, txtFzMsg);">
        <SPAN id="txtFzMsg"></td>
    </tr>
    <tr>
      <td class=th>备注<br>
        (更改原因)</td>
      <td class=ltd><textarea cols="50" rows="7" name="kpbz"><%=Rs("kp_bz")%></textarea></td>
    </Tr>
    <tr>
      <td class=th>操作</td>
      <td class=ltd><input type="submit" value=" 更改 "></td>
    </Tr>
    <input type="hidden" name="kpzrr" value="<%=Rs("kp_zrr")%>">
    <input type="hidden" name="kpinfo" value="<%=Rs("kp_item")%>">
    <input type="hidden" name="kptime" value="<%=now()%>">
    <input type="hidden" name="zrr" value="<%=strZrr%>">
    <input type="hidden" name="kpitem" value="<%=strkpitem%>">
    <input type="hidden" name="kpgzz" value="<%=strgzz%>">
    <input type="hidden" name="kpjs" value="<%=strkpjs%>">
    <input type="hidden" name="kplsh" value="<%=strclsh%>">
    <input type="hidden" name="iPage" value="<%=striPage%>">
    <input type="hidden" name="id" value="<%=Rs("id")%>">
    <input type="hidden" name="action" value="ygkpchange">
  </form>
</Table>
<%
	Rs.Close
End Function
%>
<script   language=javascript>
  function   check(e){
  var   num=e.value;
  re=/^(([1-9]\d*\.\d{0,2})|(0\.\d{0,2})|([1-9]\d*))$/;
  if(!re.test(num))
  {alert("考评分只能是数字"); document.all.kpfz.focus();}
  }
  </script>
