<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->

<%
Call ChkPageAble(11)
CurPage="������� �� ��Ӿ���/Ԥ����ʩ��"
strPage="tech"
Call FileInc(0, "js/tech.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%Call RectifyAdd()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

function RectifyAdd()
%>
	<%Call TbTopic("��Ӿ���/Ԥ����ʩ��")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<form id=frm_Rectifyadd name=frm_Rectifyadd action=Rectify_indb.asp?action=add method=post onSubmit='return checkinf();'>

	<tr>
		<th class=th height=20>��Ŀ����</td>
		<th class=th colspan="2">��Ŀ����</td>
	</tr>

	<tr>
		<td class=rtd>���β���</td>
		<td class=ltd colspan="2"><input type=text name="zrbm" size=15>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*</td>
	</tr>
	
	<tr>
		<td class=rtd>���</td>
		<td class=ltd colspan="2"><input type=text name="bh" size=15></td>
	</tr>
	
	<tr>
		<td class=rtd>������Ϣ����</td>
		<td class=ltd colspan="2"><input type=text name="xxbm" size=15></td>
	</tr>
		
	<tr>
		<td class=rtd>��Ϣ��������</td>
		<td class=ltd colspan="2">
		<script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script>		</td>
	</tr>
			
	<tr>
		<td class=rtd>���ϸ�/Ǳ��<br>���ϸ�����</td>
		<td class=ltd colspan="2"><textarea name="bhgnr" cols="75" rows="7"></textarea>&nbsp;&nbsp;&nbsp;*</td>
	</tr>
	
	<tr>
		<td class=rtd rowspan="2" >����/Ԥ��<br>��ʩҪ��</td>
		<td class=ltd colspan="2"><textarea name="yfcsyq" cols="75" rows="7"></textarea></td>
	</tr>
	
	<tr>
		<td class=ltd>���ޣ�
		<script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='qxsj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script>		</td>
		<td class=ltd>�Ƿ�����<input type="radio" value="V1" checked name="ps">��&nbsp;&nbsp; 
		<input type="radio" name="ps" value="V2">��</td>
	</tr>

	<tr>
		<td class=rtd>�������<br>ԭ�����</td>
		<td class=ltd colspan="2"><textarea name="yyfx" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd>����/Ԥ����ʩ</td>
		<td class=ltd colspan="2"><textarea name="jzcs" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd>��ʵ���</td>
		<td class=ltd colspan="2"><textarea name="lsqk" cols="75" rows="7"></textarea></td>
	</tr>
	
	<tr>
		<td class=rtd>��֤����</td>
		<td class=ltd colspan="2"><textarea name="yzjl" cols="75" rows="7"></textarea></td>
	</tr>

	<tr><td class=ctd colspan=3><input type=submit value=" �� ȷ �� �� "></td></tr>
	</form>
	</table>
<%
end function	
%>