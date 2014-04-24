<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/calendar.asp"-->
<%
Call ChkPageAble(11)
CurPage="问题分析 → 更改外部质量信息"
strPage="tech"
Call FileInc(0, "js/tech.js")
xjweb.header()
Call TopTable()
Dim iid
iid=0
If IsNumeric(Trim(Request("id"))) Then iid=CLng(Trim(Request("id")))
If iid=0 Then
	Call JsAlert("请从正规入口进入!谢谢!","quality_list.asp")
Else
	Call Main()
End If
Call BottomTable()
xjweb.footer()
closeObj()

Function Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd height=300>
			<%
				If iid=0 Then Exit Function
				strSql="select * from [quality] where id="&iid&""
				Set Rs=xjweb.Exec(strSql, 1)
				If Rs.Eof Or Rs.Bof Then Rs.Close : Exit Function
				Call qualityChange(Rs)
				Rs.Close
				Response.Write(XjLine(10,"100%",""))
			%>
		</Td></Tr>
	</Table>
<%
End Function

Function qualityChange(Rs)
	Call TbTopic("修改外部质量信息")
%>
	<table class=xtable cellspacing=0 cellpadding=3 width="80%">
	<form id=frm_qualityadd name=frm_qualityadd action=quality_indb.asp?action=change method=post onSubmit='return checkinf();'>
		
	<tr>
		<th class=th height=25>项目名称</td>
		<th class=th>项目内容</td>
	</tr>

	<tr>
		<td class=rtd>客户名称</td>
		<td class=ltd><input type=text name="khmc" size=15 value="<%=Rs("khmc")%>"></td>
	</tr>
	
	<tr>
		<td class=rtd>联系人</td>
		<td class=ltd><input type=text name="lxr" size=15 value="<%=Rs("lxr")%>"></td>
	</tr>
	
	<tr>
		<td class=rtd>联系电话</td>
		<td class=ltd><input type=text name="lxdh" size=15 value="<%=Rs("lxdh")%>"></td>
	</tr>
	
	<tr>
		<td class=rtd>合同号</td>
		<td class=ltd><input type=text name="hth" size=15 value="<%=Rs("hth")%>"></td>
	</tr>

	<tr>
		<td class=rtd>工作令号</td>
		<td class=ltd><input type=text name="gzlh" size=15 value="<%=Rs("gzlh")%>"></td>
	</tr>
	
	<tr>
		<td class=rtd>接收时间</td>
		<td class=ltd>
		<script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='jssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;原接收时间：<%=rs("jssj")%>		</td>
	</tr>
			
	<tr>
		<td class=rtd>责任人</td>
		<td class=ltd><input type=text name="zrr" size=15 value="<%=Rs("zrr")%>"></td>
	</tr>

	<tr>
		<td class=rtd>主要问题</td>
		<td class=ltd><textarea name="zywt" cols="75" rows="7"><%=rs("zywt")%></textarea></td>
	</tr>
	<tr>
	</tr>
	
	<tr>
		<td class=rtd>应急措施</td>
		<td class=ltd><textarea name="yjcs" cols="75" rows="7"><%=rs("yjcs")%></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr>
		<td class=rtd>原因分析</td>
		<td class=ltd><textarea name="yyfx" cols="75" rows="7"><%=Rs("yyfx")%></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr>
		<td class=rtd>纠正措施</td>
		<td class=ltd><textarea name="jzcs" cols="75" rows="7"><%=Rs("jzcs")%></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr>
		<td class=rtd>落实情况</td>
		<td class=ltd><textarea name="lsqk" cols="75" rows="7"><%=Rs("lsqk")%></textarea></td>
	</tr>
	<tr>
	</tr>
	
	<tr>
		<td class=rtd>验证结论</td>
		<td class=ltd><textarea name="yzjl" cols="75" rows="7"><%=Rs("yzjl")%></textarea></td>
	</tr>
	<tr>
	</tr>

	<tr><td class=ctd colspan=2><input type=submit value=" · 确 定 · "></td></tr>
	<input type="hidden" name="id" value=<%=Rs("id")%>>
	</form>
	</table>
<%
end function	
%>