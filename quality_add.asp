<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->

<%
Call ChkPageAble(11)
CurPage="������� �� ����ⲿ������Ϣ"
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
			<%Call qualityAdd()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

function qualityAdd()
%>
	<%Call TbTopic("����ⲿ������Ϣ")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<form id=frm_qualityadd name=frm_qualityadd action=quality_indb.asp?action=add method=post onSubmit='return checkinf();'>

	<tr>
		<th class=th height=25>��Ŀ����</td>
		<th class=th>��Ŀ����</td>
	</tr>

	<tr>
		<td class=rtd>�ͻ�����</td>
		<td class=ltd><input type=text name="khmc" size=15>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*</td>
	</tr>
	
	<tr>
		<td class=rtd>��ϵ��</td>
		<td class=ltd><input type=text name="lxr" size=15></td>
	</tr>
	
	<tr>
		<td class=rtd>��ϵ�绰</td>
		<td class=ltd><input type=text name="lxdh" size=15></td>
	</tr>
	
	<tr>
		<td class=rtd>��ͬ��</td>
		<td class=ltd><input type=text name="hth" size=15>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*</td>
	</tr>

	<tr>
		<td class=rtd>�������</td>
		<td class=ltd><input type=text name="gzlh" size=15></td>
	</tr>
	
	<tr>
		<td class=rtd>����ʱ��</td>
		<td class=ltd>
		<script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script>		</td>
	</tr>
			
	<tr>
		<td class=rtd>������</td>
		<td class=ltd><input type=text name="zrr" size=15>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*</td>
	</tr>

	<tr>
		<td class=rtd>��Ҫ����</td>
		<td class=ltd><textarea name="zywt" cols="75" rows="7"></textarea>&nbsp;&nbsp;&nbsp;*</td>
	</tr>
	<tr>
	</tr>
	
	<tr>
		<td class=rtd>Ӧ����ʩ</td>
		<td class=ltd><textarea name="yjcs" cols="75" rows="7"></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr>
		<td class=rtd>ԭ�����</td>
		<td class=ltd><textarea name="yyfx" cols="75" rows="7"></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr>
		<td class=rtd>������ʩ</td>
		<td class=ltd><textarea name="jzcs" cols="75" rows="7"></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr>
		<td class=rtd>��ʵ���</td>
		<td class=ltd><textarea name="lsqk" cols="75" rows="7"></textarea></td>
	</tr>
	<tr>
	</tr>
	
	<tr>
		<td class=rtd>��֤����</td>
		<td class=ltd><textarea name="yzjl" cols="75" rows="7"></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr><td class=ctd colspan=2><input type=submit value=" �� ȷ �� �� "></td></tr>
	</form>
	</table>
<%
end function	
%>