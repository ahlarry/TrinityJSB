<%
function mtask_display_few(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>����ģ�߳�����ģ���������</font>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4><b>����ͬ��Ϣ��</b></td>
		</tr>

		<tr>
			<td class=td_rblue width="15%">������</td>
			<td class=td_lblue width="35%"><%=rs("ddh")%></td>
			<td class=td_rblue width="15%">��ˮ��</td>
			<td class=td_lblue width="*"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
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
			<td class=td_rblue>��������</td>
			<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
			<td class=td_rblue>�ƻ�����ʱ��</td>
			<td class=td_lblue><%=rs("jhjssj")%></td>
		</tr>
	</table>
<%
end function

function mtask_display_much(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>����ģ�߳�����ģ���������</font>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4><b>��ͬ��Ϣ</b></td>
		</tr>

		<tr>
			<td class=td_rblue width="13%">������</td>
			<td class=td_lblue width="37%"><%=rs("ddh")%></td>
			<td class=td_rblue width="13%">��ˮ��</td>
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
			<td class=td_rblue>ģ�߲���</td>
			<td class=td_lblue><%=rs("mjcl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>�豸����</td>
			<td class=td_lblue><%=rs("sbcj")%></td>
			<td class=td_rblue>ˮ��ͷ����</td>
			<td class=td_lblue><%=rs("sjtsl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>�������ͺ�</td>
			<td class=td_lblue><%=rs("jcjxh")%></td>
			<td class=td_rblue>����ͷ����</td>
			<td class=td_lblue><%=rs("qjtsl")%></td>
			
		</tr>
		<tr>
			<td class=td_rblue>����Ȱ�</td>
			<td class=td_lblue><%if rs("pjrb") then%>��<%else%>��<%end if%></td>
			<td class=td_rblue>ǻ��</td>
			<td class=td_lblue>
				<%if rs("qs")=1 then%>��ǻ<%end if%>
				<%if rs("qs")=2 then%>˫ǻ<%end if%>
				<%if rs("qs")=3 then%>��ǻ<%end if%>
				<%if rs("qs")=4 then%>��ǻ<%end if%>
				<%if rs("qs")=5 then%>��ǻ<%end if%>
				<%if rs("qs")=6 then%>��ǻ<%end if%>
				<%if rs("qs")=7 then%>��ǻ<%end if%>
				<%if rs("qs")=8 then%>��ǻ<%end if%>
			</td>
		</tr>
		
		<tr>
			<td class=td_rblue>���Ȱ���Ϣ</td>
			<td class=td_lblue>����:<%=rs("jrbxs")%>	 ����:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
			<td class=td_rblue>ǣ���ٶ�</td>
			<td class=td_lblue><%=rs("qysd")%> ��/��(m/min)</td>
		</tr>

		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>ģ����Ϣ</b></td>
		</tr>

		<tr>
			<td class=td_rblue>ģͷ�ṹ</td>
			<td class=td_lblue><%=rs("mtjg")%>&nbsp;</td>
			<td class=td_rblue>��������</td>
			<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
		</tr>

		<tr>
			<td class=td_rblue>���ͽṹ</td>
			<td class=td_lblue><%=rs("dxjg")%>&nbsp;</td>
			<td class=td_rblue>ģͷ���ӳߴ�</td>
			<td class=td_lblue><%=rs("mtljcc")%></td>
		</tr>

		<tr>
			<td class=td_rblue>ˮ��ṹ</td>
			<td class=td_lblue><%=rs("sxjg")%>&nbsp;</td>
			<td class=td_rblue>�ȵ�ż���</td>
			<td class=td_lblue><%=rs("rdogg")%></td>
		</tr>


		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>������Ϣ</b></td>
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
			<td class=td_lblue height=25 colspan=4> <b>������Ϣ</b></td>
		</tr>

		<tr>
			<td class=td_rblue >�����¼</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("psjl"))%></td>
		</tr>

		<tr>
			<td class=td_rblue >��ע</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("bz"))%></td>
		</tr>

		<tr>
			<td class=td_rblue>�ƻ�����ʱ��</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
			<td class=td_rblue>ʵ�ʽ���ʱ��;</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
		</tr>

		<tr>
			<td class=td_rblue>�鳤</td>
			<td class=td_lblue><%=rs("zz")%></td>
			<td class=td_rblue>��������</td>
			<td class=td_lblue><%=rs("jsdb")%></td>
		</tr>
	</table>
<%
end function
 
function mtask_display_all(rs)
%>
	<font style=font-size:20px;font-weight:bold;text-align:center;>����ģ�߳�����ģ���������</font>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4><b>��ͬ��Ϣ</b></td>
		</tr>

		<tr>
			<td class=td_rblue width="13%">������</td>
			<td class=td_lblue width="37%"><%=rs("ddh")%></td>
			<td class=td_rblue width="13%">��ˮ��</td>
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
			<td class=td_rblue>ģ�߲���</td>
			<td class=td_lblue><%=rs("mjcl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>�豸����</td>
			<td class=td_lblue><%=rs("sbcj")%></td>
			<td class=td_rblue>ˮ��ͷ����</td>
			<td class=td_lblue><%=rs("sjtsl")%></td>
		</tr>

		<tr>
			<td class=td_rblue>�������ͺ�</td>
			<td class=td_lblue><%=rs("jcjxh")%></td>
			<td class=td_rblue>����ͷ����</td>
			<td class=td_lblue><%=rs("qjtsl")%></td>
			
		</tr>
		<tr>
			<td class=td_rblue>����Ȱ�</td>
			<td class=td_lblue><%if rs("pjrb") then%>��<%else%>��<%end if%></td>
			<td class=td_rblue>ǻ��</td>
			<td class=td_lblue>
				<%if rs("qs")=1 then%>��ǻ<%end if%>
				<%if rs("qs")=2 then%>˫ǻ<%end if%>
				<%if rs("qs")=3 then%>��ǻ<%end if%>
				<%if rs("qs")=4 then%>��ǻ<%end if%>
				<%if rs("qs")=5 then%>��ǻ<%end if%>
				<%if rs("qs")=6 then%>��ǻ<%end if%>
				<%if rs("qs")=7 then%>��ǻ<%end if%>
				<%if rs("qs")=8 then%>��ǻ<%end if%>
			</td>
		</tr>
		
		<tr>
			<td class=td_rblue>���Ȱ���Ϣ</td>
			<td class=td_lblue>����:<%=rs("jrbxs")%>	 ����:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
			<td class=td_rblue>ǣ���ٶ�</td>
			<td class=td_lblue><%=rs("qysd")%> ��/��(m/min)</td>
		</tr>

		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>ģ����Ϣ</b></td>
		</tr>

		<tr>
			<td class=td_rblue>ģͷ�ṹ</td>
			<td class=td_lblue><%=rs("mtjg")%>&nbsp;</td>
			<td class=td_rblue>��������</td>
			<td class=td_lblue><%=rs("mjxx") & rs("rwlr")%></td>
		</tr>

		<tr>
			<td class=td_rblue>���ͽṹ</td>
			<td class=td_lblue><%=rs("dxjg")%>&nbsp;</td>
			<td class=td_rblue>ģͷ���ӳߴ�</td>
			<td class=td_lblue><%=rs("mtljcc")%></td>
		</tr>

		<tr>
			<td class=td_rblue>ˮ��ṹ</td>
			<td class=td_lblue><%=rs("sxjg")%>&nbsp;</td>
			<td class=td_rblue>�ȵ�ż���</td>
			<td class=td_lblue><%=rs("rdogg")%></td>
		</tr>


		<tr bgcolor="#DDDDDD">
			<td class=td_lblue height=25 colspan=4> <b>������Ϣ</b></td>
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
			<td class=td_lblue height=25 colspan=4> <b>������Ϣ</b></td>
		</tr>

		<tr>
			<td class=td_rblue >�����¼</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("psjl"))%></td>
		</tr>

		<tr>
			<td class=td_rblue >��ע</td>
			<td class=td_lblue colspan=3 height=200 valign=top><%=xujian_ims.htmltocode(rs("bz"))%></td>
		</tr>

		<tr>
			<td class=td_rblue>�ƻ�����ʱ��</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
			<td class=td_rblue>ʵ�ʽ���ʱ��;</td>
			<td class=td_lblue><%=xujian_date(rs("jhjssj"),1)%></td>
		</tr>

		<tr>
			<td class=td_rblue>�鳤</td>
			<td class=td_lblue><%=rs("zz")%></td>
			<td class=td_rblue>��������</td>
			<td class=td_lblue><%=rs("jsdb")%></td>
		</tr>
	</table>
<%
	response.write(hline(5, "100%", ""))
	call mtask_display_user(rs)
	response.write(hline(5, "100%", ""))
	call mtask_display_user2(rs)
end function

function mtask_display_user(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr>
		<%
			select case rs("mjxx") & rs("rwlr")
				case "ȫ�����"
				%>
					<%if (not isnull(rs("mtjgr"))) and (not isnull(rs("dxjgr"))) and (rs("mtjgr")=rs("dxjgr")) then%>
						<td class=td_blue width="10%" rowspan=2>ģ�߽ṹ</td>
					<%else%>
						<td class=td_blue width="10%">ģͷ�ṹ</td>
					<%end if%>
					<%if rs("mtjgr")=rs("dxjgr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtjgr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtjgr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ�����</td>
					<%else%>
						<td class=td_blue width="10%">ģͷ���</td>
					<%end if%>
					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtsjr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ�����</td>
					<%else%>
						<td class=td_blue width="10%">ģͷ���</td>
					<%end if%>
					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtshr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ��BOM</td>
					<%else%>
						<td class=td_blue width="10%">ģͷBOM</td>
					<%end if%>
					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mtbomr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
						<td class=td_blue width="10%">���ͽṹ</td>
					<%end if%>
					<%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
						<td class=td_blue width="15%"><%=rs("dxjgr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="10%">�������</td>
					<%end if%>
					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="10%">�������</td>
					<%end if%>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="10%">����BOM</td>
					<%end if%>
					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
					<%end if%>
				<%
				case "ȫ�׸���"
				%>
					<td class=td_blue width="10%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>

					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ�߸���</td>
					<%else%>
						<td class=td_blue width="10%">ģͷ����</td>
					<%end if%>
					<%if rs("mtsjr")=rs("dxsjr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtsjr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ�����</td>
					<%else%>
						<td class=td_blue width="10%">ģͷ���</td>
					<%end if%>
					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtshr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ��BOM</td>
					<%else%>
						<td class=td_blue width="10%">ģͷBOM</td>
					<%end if%>
					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mtbomr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="10%">���͸���</td>
					<%end if%>
					<%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
						<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="10%">�������</td>
					<%end if%>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="10%">����BOM</td>
					<%end if%>
					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
					<%end if%>
				<%
				case "ȫ�׸���"
				%>
					<td class=td_blue width="10%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="15%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="10%" rowspan=2>&nbsp;</td>
					<td class=td_blue width="15%" rowspan=2>&nbsp;</td>

					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ�߸���</td>
					<%else%>
						<td class=td_blue width="10%">ģͷ����</td>
					<%end if%>
					<%if rs("mtshr")=rs("dxshr") then%>
						<td class=td_blue width="15%" rowspan=2><%=rs("mtshr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="10%" rowspan=2>ģ��BOM</td>
					<%else%>
						<td class=td_blue width="10%">ģͷBOM</td>
					<%end if%>
					<%if rs("mtbomr")=rs("dxbomr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mtbomr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="10%">���͸���</td>
					<%end if%>
					<%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
						<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="10%">����BOM</td>
					<%end if%>
					<%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
						<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
					<%end if%>
				<%
				case "ģͷ���"
				%>
					<td class=td_blue width="10%">ģͷ�ṹ</td>
					<td class=td_blue width="15%"><%=rs("mtjgr")%>&nbsp;</td>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
				<%
				case "ģͷ����"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="15%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
				<%
				case "ģͷ����"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="15%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="*"><%=rs("mtbomr")%>&nbsp;</td>
				<%
				case "�������"
				%>
					<td class=td_blue width="10%">���ͽṹ</td>
					<td class=td_blue width="15%"><%=rs("dxjgr")%>&nbsp;</td>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
				<%
				case "���͸���"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">���͸���</td>
					<td class=td_blue width="15%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
				<%
				case "���͸���"
				%>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">&nbsp;</td>
					<td class=td_blue width="15%">&nbsp;</td>
					<td class=td_blue width="10%">���͸���</td>
					<td class=td_blue width="15%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="*"><%=rs("dxbomr")%>&nbsp;</td>
				<%
				case else
				response.write(rs("mjxx") & rs("rwlr"))
			end select
		%>
		</tr>
	</table>
<%
end function

function mtask_display_user2(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<tr>
		<%
			select case rs("mjxx")
				case "ȫ��"
				%>
					<%if (not isnull(rs("mttsdr"))) and (not isnull(rs("dxtsdr"))) and (rs("mttsdr")=rs("dxtsdr")) then%>
						<td class=td_blue width="15%" rowspan=2>ģ�ߵ��Ե�</td>
					<%else%>
						<td class=td_blue width="15%">ģͷ���Ե�</td>
					<%end if%>
					<%if rs("mttsdr")=rs("dxtsdr") then%>
						<td class=td_blue width="18%" rowspan=2><%=rs("mttsdr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="18%"><%=rs("mttsdr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mttsr")=rs("dxtsr") then%>
						<td class=td_blue width="15%" rowspan=2>ģ�ߵ���</td>
					<%else%>
						<td class=td_blue width="15%">ģͷ����</td>
					<%end if%>
					<%if rs("mttsr")=rs("dxtsr") then%>
						<td class=td_blue width="18%" rowspan=2><%=rs("mttsr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="18%"><%=rs("mttsr")%>&nbsp;</td>
					<%end if%>

					<%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
						<td class=td_blue width="15%" rowspan=2>ģ�ߵ�����Ϣ����</td>
					<%else%>
						<td class=td_blue width="15%">ģͷ������Ϣ����</td>
					<%end if%>
					<%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
						<td class=td_blue width="*" rowspan=2><%=rs("mttsxxzlr")%>&nbsp;</td>
					<%else%>
						<td class=td_blue width="*"><%=rs("mttsxxzlr")%>&nbsp;</td>
					<%end if%>
				</tr>
				<tr>
					<%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
						<td class=td_blue width="15%">���͵��Ե�</td>
					<%end if%>
					<%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
						<td class=td_blue width="18%"><%=rs("dxtsdr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
						<td class=td_blue width="15%">���͵���</td>
					<%end if%>
					<%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
						<td class=td_blue width="18%"><%=rs("dxtsr")%>&nbsp;</td>
					<%end if%>

					<%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
						<td class=td_blue width="15%">���͵�����Ϣ����</td>
					<%end if%>
					<%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
						<td class=td_blue width="*"><%=rs("dxtsxxzlr")%>&nbsp;</td>
					<%end if%>
				<%
				case "ģͷ"
				%>
					<td class=td_blue width="15%">ģͷ���Ե�</td>
					<td class=td_blue width="18%"><%=rs("mttsdr")%>&nbsp;</td>
					<td class=td_blue width="15%">ģͷ����</td>
					<td class=td_blue width="18%"><%=rs("mttsr")%>&nbsp;</td>
					<td class=td_blue width="15%">ģͷ������Ϣ����</td>
					<td class=td_blue width="*"><%=rs("mttsxxzlr")%>&nbsp;</td>
				<%
				case "����"
				%>
					<td class=td_blue width="15%">���͵��Ե�</td>
					<td class=td_blue width="18%"><%=rs("dxtsdr")%>&nbsp;</td>
					<td class=td_blue width="15%">���͵���</td>
					<td class=td_blue width="18%"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=td_blue width="15%">���͵�����Ϣ����</td>
					<td class=td_blue width="*"><%=rs("dxtsxxzlr")%>&nbsp;</td>
				<%
				case else
				response.write(rs("mjxx"))
			end select
		%>
		</tr>
	</table>
<%
end function


function mtask_display_user_all(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
	<%select case rs("mjxx") & rs("rwlr")%>
			<%case "ȫ�����"%>
				<tr>
					<td class=td_blue width="10%">ģͷ�ṹ</td>
					<td class=td_blue width="10%"><%=rs("mtjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ�ṹ��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ�ṹ����</td>
					<td class=td_blue width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��ƿ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��ƽ���</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˽���</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM����</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���ͽṹ</td>
					<td class=td_blue width="10%"><%=rs("dxjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">���ͽṹ��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">���ͽṹ����</td>
					<td class=td_blue width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">������ƿ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">������ƽ���</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">������˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">������˽���</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM����</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "ģͷ���"%>
				<tr>
					<td class=td_blue width="10%">ģͷ�ṹ</td>
					<td class=td_blue width="10%"><%=rs("mtjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ�ṹ��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ�ṹ����</td>
					<td class=td_blue width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��ƿ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��ƽ���</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˽���</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM����</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
			<%case "�������"%>
				<tr>
					<td class=td_blue width="10%">���ͽṹ</td>
					<td class=td_blue width="10%"><%=rs("dxjgr")%>&nbsp;</td>
					<td class=td_blue width="20%">���ͽṹ��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxjgks")%>&nbsp;</td>
					<td class=td_blue width="20%">���ͽṹ����</td>
					<td class=td_blue width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">������ƿ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">������ƽ���</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">������˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">������˽���</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM����</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "ȫ�׸���"%>
				<tr>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ŀ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ľ���</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˽���</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM����</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���͸���</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸��Ŀ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸��Ľ���</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">������˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">������˽���</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM����</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "ģͷ����"%>
				<tr>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="10%"><%=rs("mtsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ŀ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ľ���</td>
					<td class=td_blue width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ���</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ��˽���</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM����</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
			<%case "���͸���"%>
				<tr>
					<td class=td_blue width="10%">���͸���</td>
					<td class=td_blue width="10%"><%=rs("dxsjr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸��Ŀ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxsjks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸��Ľ���</td>
					<td class=td_blue width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">�������</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">������˿�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">������˽���</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM����</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "ȫ�׸���"%>
				<tr>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���鿪ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ�������</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM����</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���͸���</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸��鿪ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸������</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM����</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
			<%case "ģͷ����"%>
				<tr>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="10%"><%=rs("mtshr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���鿪ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtshks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ�������</td>
					<td class=td_blue width="20%"><%=rs("mtshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷBOM</td>
					<td class=td_blue width="10%"><%=rs("mtbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("mtbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷBOM����</td>
					<td class=td_blue width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
				</tr>
			<%case "���͸���"%>
				<tr>
					<td class=td_blue width="10%">���͸���</td>
					<td class=td_blue width="10%"><%=rs("dxshr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸��鿪ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxshks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͸������</td>
					<td class=td_blue width="20%"><%=rs("dxshjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">����BOM</td>
					<td class=td_blue width="10%"><%=rs("dxbomr")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM��ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxbomks")%>&nbsp;</td>
					<td class=td_blue width="20%">����BOM����</td>
					<td class=td_blue width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
				</tr>
		<%end select%>
	</table>
<%
end function

function mtask_display_user2_all(rs)
%>
	<table class=table_blue cellspacing=0 cellpadding=3 width="98%">
		<%select case rs("mjxx")%>
			<%case "ȫ��"%>
				<tr>
					<td class=td_blue width="10%">ģͷ���Ե�</td>
					<td class=td_blue width="10%"><%=rs("mttsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ե���ʼ</td>
					<td class=td_blue width="20%"><%=rs("mttsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ե�����</td>
					<td class=td_blue width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="10%"><%=rs("mttsr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Կ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mttsks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Խ���</td>
					<td class=td_blue width="20%"><%=rs("mttsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ������Ϣ����</td>
					<td class=td_blue width="10%"><%=rs("mttsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ������Ϣ����ʼ</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ������Ϣ�������</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���͵��Ե�</td>
					<td class=td_blue width="10%"><%=rs("dxtsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Ե���ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Ե�����</td>
					<td class=td_blue width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���͵���</td>
					<td class=td_blue width="10%"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Կ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxtsks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Խ���</td>
					<td class=td_blue width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���͵�����Ϣ����</td>
					<td class=td_blue width="10%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵�����Ϣ����ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵�����Ϣ�������</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
				</tr>
			<%case "ģͷ"%>
				<tr>
					<td class=td_blue width="10%">ģͷ���Ե�</td>
					<td class=td_blue width="10%"><%=rs("mttsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ե���ʼ</td>
					<td class=td_blue width="20%"><%=rs("mttsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Ե�����</td>
					<td class=td_blue width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ����</td>
					<td class=td_blue width="10%"><%=rs("mttsr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Կ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("mttsks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ���Խ���</td>
					<td class=td_blue width="20%"><%=rs("mttsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">ģͷ������Ϣ����</td>
					<td class=td_blue width="10%"><%=rs("mttsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ������Ϣ����ʼ</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">ģͷ������Ϣ�������</td>
					<td class=td_blue width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
				</tr>
			<%case "����"%>
				<tr>
					<td class=td_blue width="10%">���͵��Ե�</td>
					<td class=td_blue width="10%"><%=rs("dxtsdr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Ե���ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Ե�����</td>
					<td class=td_blue width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���͵���</td>
					<td class=td_blue width="10%"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Կ�ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxtsks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵��Խ���</td>
					<td class=td_blue width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
				</tr>
				<tr>
					<td class=td_blue width="10%">���͵�����Ϣ����</td>
					<td class=td_blue width="10%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵�����Ϣ����ʼ</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
					<td class=td_blue width="20%">���͵�����Ϣ�������</td>
					<td class=td_blue width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
				</tr>

			<%
				case else
				response.write(rs("mjxx"))
			end select
			%>
		</tr>
	</table>
<%
end function
%>