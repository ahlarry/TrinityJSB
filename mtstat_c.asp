<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<style type="text/css">
<!--
.btn_2k3 {
	BORDER-RIGHT: #002D96 1px solid;
	PADDING-RIGHT: 2px;
	BORDER-TOP: #002D96 1px solid;
	PADDING-LEFT: 2px;
	FONT-SIZE: 12px;
FILTER: progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#FFFFFF, EndColorStr=#9DBCEA);
	BORDER-LEFT: #002D96 1px solid;
	CURSOR: hand;
	COLOR: black;
	PADDING-TOP: 2px;
	BORDER-BOTTOM: #002D96 1px solid
}
-->
</style>
<HTML>
<HEAD>
<TITLE>����Ա���Կ���ϵ���޸�</TITLE>
</head>
<BODY bgColor=#95db98 topmargin=15 leftmargin=15 >
<object id="WebBrowser" width=0 height=0 classid="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2">
</object>
<%
Call ChkPageAble(3)
	dim action
	action=request("action")
	select case action
		case "chg_indb"
			call chg_indb()
		case else
			call chg_dis()
	end select
%>
</body>
</html>
<%
Function chg_dis()
Dim iid, ilsh, ixs, stryfz
iid=Request("id")
If Not IsNumeric(iid) Then Call JsAlert("�����ȷ��ڽ���!","index.asp")
set rs=server.createobject("adodb.recordset")
strSql="select * from [mantime] where ID=" & Clng(iid)
rs.open strSql,conn,1,1
ilsh=Rs("lsh")
If NullToNum(Rs("rwfz"))=0 Then
	stryfz=Rs("fz")
else
	stryfz=Rs("rwfz")
End If
If NullToNum(Rs("jc"))=0 Then
	ixs=1
else
	ixs=Rs("jc")
End If
%>
<form name="form1" action="?action=chg_indb"  method="post" onSubmit="return chkinf();">
  <table width=100% border="0" cellpadding="0" cellspacing="2">
    <tr>
      <td><FIELDSET align=left>
          <LEGEND align=left>�޸�&nbsp;<font color="#ff0000"><strong><%=ilsh%></strong></font>&nbsp;���Ժϸ�ʵ�ʵ÷�</LEGEND>
          <TABLE border="0" cellpadding="0" cellspacing="3">
            <TR>
              <TD align="left" width="48%">ԭ��ֵ��</td>
              <TD align="left"><%=stryfz%></td>
            </tr>
            <TR>
              <TD align="left" width="48%">�ַ�ֵ��</td>
              <TD align="left"><%=Rs("fz")%></td>
            </tr>
            <tr>
              <TD align="left">ϵ&nbsp;&nbsp;����</td>
              <TD align="left"><input name="sjf" type="text" style="width:60px;ime-mode:disabled" onKeyPress="checkIsFloat(this.value);" value="<%=ixs%>" maxlength="5" onpaste="return !clipboardData.getData('text').match(/\D/)" />
            </TR>
            <tr>
              <td height="1" colspan="2"><input type="hidden" name="lsh" value=<%=ilsh%>></td>
            </tr>
          </TABLE>
        </fieldset></td>
    </tr>
    <tr>
      <td height="1" colspan="2"><input type="hidden" name="id" value=<%=iid%>></td>
    </tr>
    <tr>
      <td align="center"><input name="cmdOK" class=btn_2k3 type="submit" id="cmdOK" value="  ȷ��  ">
        <input name="cmdCancel" class=btn_2k3 type=button id="cmdCancel" onClick="window.close();" value='  ȡ��  '></td>
    </tr>
  </table>
</form>
<%Rs.Close
End Function

Function chg_indb()
	dim stryfz, str_newxs, str_lsh, TmpSql, TmpRs
	str_newxs="" : str_lsh=""

	str_newxs=NullToNum(Trim(Request("sjf")))
	str_lsh=Trim(Request("lsh"))

	if str_lsh="" then
		Call JsAlert("��ˮ�Ų�ȷ��,�޷�ȷ������ĵ�������!\n���������ڽ���!","")
	end if

	TmpSql="select * from [mantime] where rwlr like '%���Ժϸ�(%' and lsh='"&str_lsh&"'"
	Set TmpRs=Server.CreateObject("adodb.recordset")
	TmpRs.open TmpSql,conn,1,3
	Do while not TmpRs.Eof
		stryfz=NullToNum(TmpRs("rwfz"))
		if stryfz=0 Then
			TmpRs("rwfz")=TmpRs("fz")
			TmpRs("fz")=Round(TmpRs("fz")*str_newxs,1)
		else
			TmpRs("fz")=Round(stryfz*str_newxs,1)
		End If
		TmpRs("jc")=str_newxs
		TmpRs("xgsj")=now()
		TmpRs("bz")=session("userName")&"�޸�ϵ��"
	TmpRs.update
	TmpRs.movenext
	Loop
	TmpRs.close
	Response.Write("<script language='JavaScript'>document.all.WebBrowser.ExecWB(45,1)</script>")
End Function
%>
<script language="JavaScript">
function chkinf(){
  var strurl=document.form1.sjf.value;
  if (strurl=="")
  {
  	alert("ϵ������Ϊ�գ�");
	document.form1.sjf.focus();
	return false;
  }
  else
  {
    window.returnValue = 2;
    window.opener=null;
    window.close();
  }
}

function checkIsFloat(arg){
 var nc=event.keyCode;
 if((nc>=48)   &&   (nc<=57)   ){
 }else   if(nc==46){
     for(var   i=0;i<arg.length;i++){
         if(arg.charAt(i)=='.'){
                     event.keyCode=0;   return;
         }
     }
 }else{
     event.keyCode=0;return;
 }
 }
</script>
