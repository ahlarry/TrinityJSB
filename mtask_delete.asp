<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'11:45 2007-1-8-����һ
Call ChkPageAble(3)
CurPage="������� �� ɾ��������"
strPage="mtask"
'Call FileInc(0, "js/mtask.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
	Dim iid
	iid=Request.Form("id")
	If iid<>"" Then
		Call mtask_db_delete()
	Else
	%>
		<Table class=xtable cellspacing=0 cellpadding=2 width="<%=web_info(8)%>">
			<Tr><Td class=ctd>
				<%Call SearchLsh()%>
			</td></tr>
			<Tr><Td class=ctd height=300>
				<%Call mtaskDelete()%>
				<%Response.Write(XjLine(10,"100%",""))%>
			</Td></Tr>
		</Table>
	<%
	End If
End Sub

Function mtaskDelete()
	Dim s_lsh
	s_lsh=""
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("������Ҫɾ�����������ˮ��!") : Exit Function

	strSql="select * from [mtask] where [lsh]='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("��ˮ��Ϊ ��" & s_lsh & "�� �������鲻����!","mtask_delete.asp")
	Else
		If Not IsNull(rs("sjjssj")) Then
			Call JsAlert("��ˮ��Ϊ ��" & s_lsh & "�� ���������Ѿ����,����ɾ��!","mtask_delete.asp")
		Else
			Call mtask_Delete(Rs)
		End If
	End If
	Rs.Close
End Function

Function mtask_delete(rs)
%>
	<%Call TbTopic("ɾ����ˮ�� <font style=color:#0000FF>" &rs("lsh")&"</font> ��������")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%">
	<form action=<%=request.servervariables("script_name")%> method=post onSubmit='return confirm("������ɾ���󽫲��ܻظ�!\n��ȷ��ɾ����ˮ�� ��<%=rs("lsh")%>�� ����������?");'>

	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>����ͬ��Ϣ</b></td>
	</tr>

	<tr>
		<td class=rtd width="20%">������</td>
		<td class=ltd width="30%"><%=rs("ddh")%></td>
		<td class=rtd width="20%">��ˮ��</td>
		<td class=ltd width="*"><%=rs("lsh")%></td>
	</tr>

	<tr>
		<td class=rtd>�ͻ�����</td>
		<td class=ltd><%=rs("dwmc")%></td>
		<td class=rtd>��������</td>
		<td class=ltd><%=rs("dmmc")%></td>
	</tr>

	<tr>
		<td class=rtd>ģ��</td>
		<td class=ltd><%=rs("mh")%></td>
		<td class=rtd>�豸����</td>
		<td class=ltd><%=rs("sbcj")%></td>
	</tr>

	<tr>
		<td class=rtd>�������ͺ�</td>
		<td class=ltd><%=rs("jcjxh")%></td>
		<td class=rtd>ˮ��ͷ����</td>
		<td class=ltd><%=rs("sjtsl")%></td>
	</tr>

	<tr>
		<td class=rtd>����ͷ����</td>
		<td class=ltd><%=rs("qjtsl")%></td>
		<td class=rtd>����Ȱ�</td>
		<td class=ltd><%if rs("pjrb") then%>��<%else%>��<%end if%></td>
	</tr>

	<tr>
		<td class=rtd>���Ȱ���Ϣ</td>
		<td class=ltd>����:<%=rs("jrbxs")%>	 ����:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
		<td class=rtd>ģ�߲���</td>
		<td class=ltd><%=rs("mjcl")%></td>
	</tr>

	<tr>
		<td class=rtd>ǻ��</td>
		<td class=ltd><%=rs("qs")%>ǻ</td>
		<td class=rtd>ǣ���ٶ�</td>
		<td class=ltd><%=rs("qysd")%>��/��(m/min)</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>��ģ����Ϣ</b></td>
	</tr>

	<tr>
		<td class=rtd>��������</td>
		<td class=ltd><%=rs("mjxx") & rs("rwlr")%></td>
		<td class=rtd>ģͷ�ṹ</td>
		<td class=ltd><%=rs("mtjg")%></td>
	</tr>

	<tr>
		<td class=rtd>���ͽṹ</td>
		<td class=ltd><%=rs("dxjg")%>&nbsp;</td>
		<td class=rtd>ˮ��ṹ</td>
		<td class=ltd><%=rs("sxjg")%>&nbsp;</td>
	</tr>

	<tr>
		<td class=rtd>ģͷ���ӳߴ�</td>
		<td class=ltd><%=rs("mtljcc")%>&nbsp;</td>
		<td class=rtd>�ȵ�ż���</td>
		<td class=ltd><%=rs("rdogg")%>&nbsp;</td>
	</tr>


	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>��������Ϣ</b></td>
	</tr>

	<tr>
		<td class=rtd>�������Ӽ���ͼ</td>
		<td class=ltd><%if rs("dxljjct") then%>��<%else%>��<%end if%></td>
		<td class=rtd>�����и�</td>
		<td class=ltd><%=rs("dxqg")%></td>
	</tr>

	<tr>
		<td class=rtd>�������׶</td>
		<td class=ltd><%if rs("ztflz") then%>��<%else%>��<%end if%></td>
		<td class=rtd>������о</td>
		<td class=ltd><%if rs("ztxx") then%>��<%else%>��><%end if%></td>
	</tr>

	<tr>
		<td class=rtd>���嶨�Ϳ�</td>
		<td class=ltd><%if rs("ztdxk") then%>��<%else%>��<%end if%></td>
		<td class=rtd>&nbsp;</td>
		<td class=ltd>&nbsp;</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=ltd height=25 colspan=4> <b>��������Ϣ</b></td>
	</tr>

	<tr>
		<td class=rtd >�����¼</td>
		<td class=ltd colspan=3><%=xjweb.HtmlToCode(Rs("psjl"))%></td>
	</tr>

	<tr>
		<td class=rtd >��ע</td>
		<td class=ltd colspan=3><%=xjweb.HtmlToCode(Rs("bz"))%></td>
	</tr>

	<tr>
		<td class=rtd>�ƻ�����ʱ��</td>
		<td class=ltd><%=rs("jhjssj")%></td>
		<td class=rtd>&nbsp;</td>
		<td class=ltd>&nbsp;</td>
	</tr>

	<tr>
    <td class=rtd>�鳤</td>
    <td class=ltd>
    <%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(�ṹ)��"&rs("sjzz")&"(���)")%>
    </td>
		<td class=rtd>��������</td>
		<td class=ltd><%=rs("jsdb")%></td>
	</tr>

	<tr><td class=ctd colspan=4><input type=submit value=" �� ɾ�� �� "></td></tr>
	<input type="hidden" name=id value=<%=rs("id")%>>
	<input type="hidden" name=s_lsh value=<%=rs("lsh")%>>
	</form>
	</table>
<%
End Function		'mtask_delete()

Function mtask_db_delete()
	Dim iid, strlsh
	iid=Request.Form("id")
	strlsh=Request.Form("s_lsh")
	strSql="delete from [mtask] where [id]=" & iid
	Call xjweb.Exec(strSql, 0)
	Call JsAlert("��ˮ�� �� " & strlsh & " �� ������ɾ���ɹ�!", "mtask_delete.asp")
End Function
%>
