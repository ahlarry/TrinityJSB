<!--#include file="include/function.asp"-->
<!--#include file="inc/mtask_dbinf.asp"-->
<%
	Call ChkPageAble(4)
	pagelink=lnk_mtask
	session("pagelink")=pagelink
	web_curtitle="  �� ɾ��������"
	xujian_ims.web_title="�������" & web_curtitle
	'xujian_ims.jsfiles_inc("js/mtask.js")
	xujian_ims.web_head()
	call lsh_search()
	response.write(hline(10, "100%", ""))

	if request.form("id") <> "" then
		call mtask_db_delete()	'ɾ���������ݲ���
	else
		call main()
	end if

	response.write(hline(10, "100%", ""))
	xujian_ims.web_foot()

function main()
	dim s_lsh
	s_lsh=""
	if trim(request("s_lsh"))<>"" then s_lsh=trim(request("s_lsh"))
	if s_lsh="" then response.write(prompt("������Ҫɾ�����������ˮ��!")) : exit function

	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	set rs=xujian_ims.exec(sql,1)
	if rs.eof or rs.bof then
		response.write(prompt("��ˮ��Ϊ <b>" & s_lsh & "</b> �������鲻����!"))
	else
		call mtask_delete(rs)
	end if
	rs.close
end function

function lsh_search()
%>
	<table border=0 cellpadding=2 cellspacing=0 width="100%">
		<form action=<%=request.servervariables("script_name")%> method=get>
		<tr>
			<td>&nbsp;&nbsp;
				����ɾ�����������ˮ��:<input type=text name=s_lsh size=8 value=<%=request("s_lsh")%>>
				<input type=submit value="����">
			</td>
		</tr>
		<tr><td class=td_frame height=1></td></tr>
		</form>
	</table>
<%
end function

function mtask_delete(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>ɾ����ˮ�� <font style=color:#0000FF><%=rs("lsh")%></font>    ��������</font>
	<table class=table_xblue cellspacing=0 cellpadding=3 width="98%">
	<form id=mtask_add name=mtask_add action=<%=request.servervariables("script_name")%> method=post onSubmit='return confirm("������ɾ���󽫲��ܻظ�!\n��ȷ��ɾ����ˮ�� <%=rs("lsh")%> ����������?");'>

	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>����ͬ��Ϣ</b></td>
	</tr>

	<tr>
		<td class=td_rblue width="20%">������</td>
		<td class=td_lblue width="30%"><%=rs("ddh")%></td>
		<td class=td_rblue width="20%">��ˮ��</td>
		<td class=td_lblue width="*"><%=rs("lsh")%></td>
	</tr>

	<tr>
		<td class=td_rblue>�ͻ�����</td>
		<td class=td_lblue><%=rs("dwmc")%></td>
		<td class=td_rblue>��������</td>
		<td class=td_lblue><%=rs("dmmc")%></td>
	</tr>

	<tr>
		<td class=td_rblue>ģ��</td>
		<td class=td_lblue><%=rs("mh")%></td>
		<td class=td_rblue>�豸����</td>
		<td class=td_lblue><%=rs("sbcj")%></td>
	</tr>

	<tr>
		<td class=td_rblue>�������ͺ�</td>
		<td class=td_lblue><%=rs("jcjxh")%></td>
		<td class=td_rblue>ˮ��ͷ����</td>
		<td class=td_lblue><%=rs("sjtsl")%></td>
	</tr>

	<tr>
		<td class=td_rblue>����ͷ����</td>
		<td class=td_lblue><%=rs("qjtsl")%></td>
		<td class=td_rblue>����Ȱ�</td>
		<td class=td_lblue><%if rs("pjrb") then%>��<%else%>��<%end if%></td>
	</tr>

	<tr>
		<td class=td_rblue>���Ȱ���Ϣ</td>
		<td class=td_lblue>����:<%=rs("jrbxs")%>	 ����:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
		<td class=td_rblue>ģ�߲���</td>
		<td class=td_lblue><%=rs("mjcl")%></td>
	</tr>

	<tr>
		<td class=td_rblue>ǻ��</td>
		<td class=td_lblue><%=rs("qs")%>ǻ</td>
		<td class=td_rblue>ǣ���ٶ�</td>
		<td class=td_lblue><%=rs("qysd")%>��/��(m/min)</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>��ģ����Ϣ</b></td>
	</tr>

	<tr>
		<td class=td_rblue>��������</td>
		<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
		<td class=td_rblue>ģͷ�ṹ</td>
		<td class=td_lblue><%=rs("mtjg")%></td>
	</tr>

	<tr>
		<td class=td_rblue>���ͽṹ</td>
		<td class=td_lblue><%=rs("dxjg")%>&nbsp;</td>
		<td class=td_rblue>ˮ��ṹ</td>
		<td class=td_lblue><%=rs("sxjg")%>&nbsp;</td>
	</tr>

	<tr>
		<td class=td_rblue>ģͷ���ӳߴ�</td>
		<td class=td_lblue><%=rs("mtljcc")%>&nbsp;</td>
		<td class=td_rblue>�ȵ�ż���</td>
		<td class=td_lblue><%=rs("rdogg")%>&nbsp;</td>
	</tr>


	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>��������Ϣ</b></td>
	</tr>

	<tr>
		<td class=td_rblue>�������Ӽ���ͼ</td>
		<td class=td_lblue><%if rs("dxljjct") then%>��<%else%>��<%end if%></td>
		<td class=td_rblue>�����и�</td>
		<td class=td_lblue><%=rs("dxqg")%></td>
	</tr>

	<tr>
		<td class=td_rblue>�������׶</td>
		<td class=td_lblue><%if rs("ztflz") then%>��<%else%>��<%end if%></td>
		<td class=td_rblue>������о</td>
		<td class=td_lblue><%if rs("ztxx") then%>��<%else%>��><%end if%></td>
	</tr>

	<tr>
		<td class=td_rblue>���嶨�Ϳ�</td>
		<td class=td_lblue><%if rs("ztdxk") then%>��<%else%>��<%end if%></td>
		<td class=td_rblue>&nbsp;</td>
		<td class=td_lblue>&nbsp;</td>
	</tr>

	<tr bgcolor="#DDDDDD">
		<td class=td_lblue height=25 colspan=4> <b>��������Ϣ</b></td>
	</tr>

	<tr>
		<td class=td_rblue >��ע</td>
		<td class=td_lblue colspan=3><%=xujian_ims.htmltocode(rs("bz"))%></td>
	</tr>

	<tr>
		<td class=td_rblue>�ƻ�����ʱ��</td>
		<td class=td_lblue><%=rs("jhjssj")%></td>
		<td class=td_rblue>&nbsp;</td>
		<td class=td_lblue>&nbsp;</td>
	</tr>

	<tr>
		<td class=td_rblue>�鳤</td>
		<td class=td_lblue><%=rs("zz")%></td>
		<td class=td_rblue>��������</td>
		<td class=td_lblue><%=rs("jsdb")%></td>
	</tr>

	<tr><td class=td_cblue colspan=4><input type=submit value=" �� ɾ�� �� "></td></tr>
	<input type="hidden" name=id value=<%=rs("id")%>>
	<input type="hidden" name=s_lsh value=<%=rs("lsh")%>>
	</form>
	</table>
<%
end function		'mtask_delete()

function mtask_db_delete()
	dim iid, strtmplsh
	iid=request.form("id")
	strtmplsh=request("s_lsh")
	sql="delete from mtask_info where id=" & iid
	call xujian_ims.exec(sql, 0)
	sql="delete from mtask_flow where lsh='"&strtmplsh&"'"
	call xujian_ims.exec(sql, 0)
	err_title="������ɾ���ɹ�!"
	err_inf="��ˮ�� <b>" & strtmplsh & "</b> ������ɾ���ɹ�!|||���<a href=mtask_delete.asp>ɾ��������</a>����ɾ��������!"
	gotoprompt(1)
	response.end
end function
%>
