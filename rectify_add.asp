<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->

<%
Call ChkPageAble(11)
CurPage="问题分析 → 添加纠正/预防措施表"
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
	<%Call TbTopic("添加纠正/预防措施表")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<form id=frm_Rectifyadd name=frm_Rectifyadd action=Rectify_indb.asp?action=add method=post onSubmit='return checkinf();'>

	<tr>
		<th class=th height=20>项目名称</td>
		<th class=th colspan="2">项目内容</td>
	</tr>

	<tr>
		<td class=rtd>责任部门</td>
		<td class=ltd colspan="2"><input type=text name="zrbm" size=15>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;*</td>
	</tr>
	
	<tr>
		<td class=rtd>编号</td>
		<td class=ltd colspan="2"><input type=text name="bh" size=15></td>
	</tr>
	
	<tr>
		<td class=rtd>发出信息部门</td>
		<td class=ltd colspan="2"><input type=text name="xxbm" size=15></td>
	</tr>
		
	<tr>
		<td class=rtd>信息发出日期</td>
		<td class=ltd colspan="2">
		<script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script>		</td>
	</tr>
			
	<tr>
		<td class=rtd>不合格/潜在<br>不合格内容</td>
		<td class=ltd colspan="2"><textarea name="bhgnr" cols="75" rows="7"></textarea>&nbsp;&nbsp;&nbsp;*</td>
	</tr>
	
	<tr>
		<td class=rtd rowspan="2" >纠正/预防<br>措施要求</td>
		<td class=ltd colspan="2"><textarea name="yfcsyq" cols="75" rows="7"></textarea></td>
	</tr>
	
	<tr>
		<td class=ltd>期限：
		<script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='qxsj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script>		</td>
		<td class=ltd>是否评审：<input type="radio" value="V1" checked name="ps">是&nbsp;&nbsp; 
		<input type="radio" name="ps" value="V2">否</td>
	</tr>

	<tr>
		<td class=rtd>问题产生<br>原因分析</td>
		<td class=ltd colspan="2"><textarea name="yyfx" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd>纠正/预防措施</td>
		<td class=ltd colspan="2"><textarea name="jzcs" cols="75" rows="7"></textarea></td>
	</tr>

	<tr>
		<td class=rtd>落实情况</td>
		<td class=ltd colspan="2"><textarea name="lsqk" cols="75" rows="7"></textarea></td>
	</tr>
	
	<tr>
		<td class=rtd>验证结论</td>
		<td class=ltd colspan="2"><textarea name="yzjl" cols="75" rows="7"></textarea></td>
	</tr>

	<tr><td class=ctd colspan=3><input type=submit value=" · 确 定 · "></td></tr>
	</form>
	</table>
<%
end function	
%>