<%
'15:08 2007-1-6-������
Rem ��ʾ������Ϣ�Ĵ���(��ΪҪ�ڶദʹ�����Է��ڴ��ļ���)
Function mtask_fewinfo(rs)
%>
<%Call TbTopic("����ģ�߳�����ģ���������")%>

<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4"><b>����ͬ��Ϣ��</b></td>
  </tr>
  <tr>
    <td class="rtd" width="15%">������</td>
    <td class="ltd" width="35%"><%=rs("ddh")%></td>
    <td class="rtd" width="15%">��ˮ��</td>
    <td width="*" class="ltd"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
  </tr>
  <tr>
    <td class="rtd">�ͻ�����</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">��������</td>
    <td class="ltd"><%=rs("dmmc")%></td>
  </tr>
  <tr>
    <td class="rtd">ģ��</td>
    <td class="ltd"><%=rs("mh")%></td>
    <td class="rtd">�豸����</td>
    <td class="ltd"><%=rs("sbcj")%></td>
  </tr>
  <tr>
    <td class="rtd">��������</td>
    <td class="ltd">
    <%
    If IsNull(rs("mtrw")) and IsNull(rs("dxrw")) Then
    	Response.Write(rs("mjxx") & rs("rwlr"))
    else
    	If Rs("mtrw")<>"" Then Response.Write("ģͷ"&rs("mtrw")) End If
    	If Rs("dxrw")<>"" Then Response.Write("����"&rs("dxrw")) End If
    End If
    %>
   </td>
    <td class="rtd">�ƻ�����ʱ��</td>
    <td class="ltd"><%=xjDate(rs("jhjssj"),1)%></td>
  </tr>
</table>
<%
End Function
'��Ҫ��Ϣ
Function mtask_muchinfo(rs)
%>
<%Call TbTopic("����ģ�߳�����ģ���������")%>
<%If Chkable("1,3") Then%>
<a href="mtask_print.asp?s_lsh=<%=rs("lsh")%>">��ӡ������</a>
<%End If%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="6"><b>��ͬ��Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd" width="13%">������</td>
    <td class="ltd"><%=rs("ddh")%></td>
    <td class="rtd" width="13%">��ˮ��</td>
    <td class="ltd" width="*"><a href="SreacDwg.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="rtd" width="13%">ģ��</td>
    <td class="ltd" width="*"><%=rs("mh")%></td>
  </tr>
  <tr>
    <td class="rtd">�ͻ�����</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">��������</td>
    <td class="ltd"><%=rs("dmmc")%></td>
    <td class="rtd">ģ�߲���</td>
    <td class="ltd"><%=rs("mjcl")%></td>
  </tr>
  <tr>
    <td class="rtd">�豸����</td>
    <td class="ltd"><%=rs("sbcj")%></td>
    <td class="rtd">ˮ��ͷ����</td>
    <td class="ltd"><%=rs("sjtsl")%></td>
    <td class="rtd">����ͷ����</td>
    <td class="ltd"><%=rs("qjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">�������ͺ�</td>
    <td class="ltd"><%=rs("jcjxh")%></td>
    <td class="rtd">��������</td>
    <td class="ltd"><%=rs("jcfx")%>&nbsp;</td>
    <td class="rtd">ǣ���ٶ�</td>
    <td class="ltd"><%=rs("qysd")%> ��/��(m/min)</td>
  </tr>
  <tr>
    <td class="rtd">����Ȱ�</td>
    <td class="ltd"><%if rs("pjrb") then%>
      ��
      <%else%>
      ��
      <%end if%></td>
    <td class="rtd">���Ȱ���Ϣ</td>
    <td class="ltd">����:<%=rs("jrbxs")%> ����:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
    <td class="rtd">ǻ��</td>
    <td class="ltd"><%=rs("qs")%>ǻ</td>
  </tr>
</table>
<%
End Function

'������Ϣ
Function mtask_technicsinfo(rs)
Dim ssgjf, qbfgjf, qgjf, hgjf,gjxx
ssgjf=NullToNum(Rs("ssgj"))
qbfgjf=NullToNum(Rs("qbfgj"))
qgjf=NullToNum(Rs("qgj"))
hgjf=NullToNum(Rs("hgj"))
gjxx=""
select case ssgjf&qbfgjf&qgjf&hgjf
Case "0000"			'����08�湲���Ʒ�ģʽ
	If Rs("gjzf")>0 and Rs("gjfs")=1 Then
 		gjxx="˫ɫ����"
 	Elseif Rs("gjzf")>0 and Rs("gjfs")=2 Then
  		gjxx="ȫ��������"
  	Elseif Rs("gjfs")=3 and Rs("qhgj")=1 Then
  		gjxx="��Ӳǰ����"
  	Elseif Rs("gjfs")=3 and Rs("qhgj")=2 Then
  		gjxx="��Ӳ�󹲼�"
  	Else
  		gjxx="/"
  	End If
Case Else		'09�湲���Ʒ�ģʽ
	If ssgjf<>0 Then gjxx="˫ɫ����"
	If qbfgjf<>0 Then gjxx=gjxx &" ȫ��������"
	If qgjf<>0 Then gjxx=gjxx &" ��Ӳǰ����"
	If hgjf<>0 Then gjxx=gjxx &" ��Ӳ�󹲼�"
end select
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="6"><b>ģ����Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd"  width="13%">��������</td>
    <td class="ltd">
    <%
    If IsNull(rs("mtrw")) and IsNull(rs("dxrw")) Then
    	Response.Write(rs("mjxx") & rs("rwlr"))
    else
    	If Rs("mtrw")<>"" Then Response.Write("ģͷ"&rs("mtrw")) End If
    	If Rs("dxrw")<>"" Then Response.Write("����"&rs("dxrw")) End If
    End If
    %>
   </td>
        <td class="rtd"  width="13%">���ڵ���</td>
    <td class="ltd"  width="*"><%if rs("cnts") then%>
      ��
      <%else%>
      &nbsp;/
      <%end if%></td>
    <td class="rtd"  width="13%">�������</td>
    <% If Rs("cnts") Then%>
    <%If Not(isnull(Rs("tslb"))) Then%>
    <td class="ltd"  width="*"><a href="mtest_display.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("tslb")%></a></td>
    <%Else%>
    <td class="ltd"  width="*">&nbsp;/</td>
    <%End If%>
    <%Else%>
    <%If Rs("beit") Then%>
    <td class="ltd"  width="*">����</td>
    <%Else%>
    <td class="ltd"  width="*">&nbsp;/</td>
    <%End If%>
    <%End If%>
  </tr>
  <tr>
    <td class="rtd">ģͷ�ṹ</td>
    <td class="ltd"><%if IsNull(rs("mtjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("mtjg"))
    End if%></td>
    <td class="rtd">���ͽṹ</td>
    <td class="ltd"><%if IsNull(rs("dxjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("dxjg"))
    End if%></td>
    <td class="rtd">ˮ��ṹ</td>
    <td class="ltd"><%if IsNull(rs("sxjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("sxjg"))
    End if%></td>
  </tr>
  <tr>
    <td class="rtd">�����и�</td>
    <td class="ltd"><%if IsNull(rs("dxqg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("dxqg"))
    End if%></td>
    <td class="rtd">ģͷ���ӳߴ�</td>
    <td class="ltd"><%=rs("mtljcc")%></td>
    <td class="rtd">�ȵ�ż���</td>
    <td class="ltd"><%=rs("rdogg")%></td>
  </tr>
  <tr>
    <td class="rtd">��������</td>
    <td class="ltd"><%=Trim(gjxx)%></td>
    <td class="rtd">�������ӳߴ�</td>
    <td class="ltd"><%=rs("gjljcc")%>&nbsp;</td>
    <td class="rtd">�Ͳıں�</td>
    <td class="ltd"><%=Rs("xcbh")%>����</td>
  </tr>
  <tr>
  </tr>
  </table>
  <%Response.Write(XjLine(5,web_info(8),""))%>
  <table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8" alt="<%=DisFzInfo(Rs)%>"><b>������Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd" >�����¼</td>
    <td class="ltd" colspan="7" height="120" valign="top"><%=xjweb.HtmlToCode(rs("psjl"))%></td>
  </tr>
  <tr>
    <td class="rtd">��ע</td>
    <td class="ltd" colspan="7" height="150" valign="top"><%=xjweb.HtmlToCode(rs("bz"))%></td>
  </tr>
  <tr>
    <td class="rtd">�ƻ���ʼ</td>
    <%If rs("jhkssj")<>"" Then%>
    <td class="ltd"><%=XjDate(rs("jhkssj"),3)%></td>
    <%else%>
    <td class="ltd" >&nbsp;/</td>
    <%End If%>
    <td class="rtd">�ƻ��ṹ����</td>
    <td class="ltd" width="12%"><%=XjDate(rs("jhjgsj"),3)%></td>
    <td class="rtd">�ƻ�ȫ�׽���</td>
    <td class="ltd"><%=XjDate(rs("jhjssj"),3)%></td>
    <td class="rtd">ʵ�ʽ���</td>
    <td class="ltd" width="12%"><%=XjDate(rs("sjjssj"),3)%></td>
  </tr>
  <tr>
    <td class="rtd">�鳤</td>
    <td colspan="3" class="ltd"><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(�ṹ)��"&rs("sjzz")&"(���)")%></td>
    <td class="rtd">��������</td>
    <td colspan="3" class="ltd"><%=rs("jsdb")%></td>
  </tr>
</table>
<% End Function
'ȫ����Ϣ
Function mtask_allinfo(rs)
%>
<%Call TbTopic("����ģ�߳�����ģ���������")%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4"><b>��ͬ��Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd" width="13%">������</td>
    <td class="ltd" width="37%"><%=rs("ddh")%></td>
    <td class="rtd" width="13%">��ˮ��</td>
    <td class="ltd" width="*"><%=rs("lsh")%></td>
  </tr>
  <tr>
    <td class="rtd">�ͻ�����</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">��������</td>
    <td class="ltd"><%=rs("dmmc")%></td>
  </tr>
  <tr>
    <td class="rtd">ģ��</td>
    <td class="ltd"><%=rs("mh")%></td>
    <td class="rtd">ģ�߲���</td>
    <td class="ltd"><%=rs("mjcl")%></td>
  </tr>
  <tr>
    <td class="rtd">�豸����</td>
    <td class="ltd"><%=rs("sbcj")%></td>
    <td class="rtd">ˮ��ͷ����</td>
    <td class="ltd"><%=rs("sjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">�������ͺ�</td>
    <td class="ltd"><%=rs("jcjxh")%></td>
    <td class="rtd">����ͷ����</td>
    <td class="ltd"><%=rs("qjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">����Ȱ�</td>
    <td class="ltd"><%if rs("pjrb") then%>
      ��
      <%else%>
      ��
      <%end if%></td>
    <td class="rtd">ǻ��</td>
    <td class="ltd"><%=rs("qs")%>ǻ</td>
  </tr>
  <tr>
    <td class="rtd">���Ȱ���Ϣ</td>
    <td class="ltd">����:<%=rs("jrbxs")%> ����:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
    <td class="rtd">ǣ���ٶ�</td>
    <td class="ltd"><%=rs("qysd")%> ��/��(m/min)</td>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4"><b>ģ����Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd">�����и�</td>
    <td class="ltd"><%=rs("dxqg")%>&nbsp;</td>
    <td class="rtd">��������</td>
    <td class="ltd">
    <%
    If IsNull(rs("mtrw")) and IsNull(rs("dxrw")) Then
    	Response.Write(rs("mjxx") & rs("rwlr"))
    else
    	If Rs("mtrw")<>"" Then Response.Write("ģͷ"&rs("mtrw")) End If
    	If Rs("dxrw")<>"" Then Response.Write("����"&rs("dxrw")) End If
    End If
    %>
   </td>
  </tr>
  <tr>
    <td class="rtd">���ͽṹ</td>
    <td class="ltd"><%=rs("dxjg")%>&nbsp;</td>
    <td class="rtd">ģͷ���ӳߴ�</td>
    <td class="ltd"><%=rs("mtljcc")%></td>
  </tr>
  <tr>
    <td class="rtd">ˮ��ṹ</td>
    <td class="ltd"><%=rs("sxjg")%>&nbsp;</td>
    <td class="rtd">�ȵ�ż���</td>
    <td class="ltd"><%=rs("rdogg")%></td>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="4" alt="<%=DisFzInfo(Rs)%>"><b>������Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd" >�����¼</td>
    <td class="ltd" colspan="3" height="120" valign="top"><%=xjweb.HtmlToCode(rs("psjl"))%></td>
  </tr>
  <tr>
    <td class="rtd" >��ע</td>
    <td class="ltd" colspan="3" height="150" valign="top"><%=xjweb.HtmlToCode(rs("bz"))%></td>
  </tr>
  <tr>
    <td class="rtd">�ƻ�����ʱ��</td>
    <td class="ltd"><%=XjDate(rs("jhjssj"),1)%></td>
    <td class="rtd">ʵ�ʽ���ʱ��;</td>
    <td class="ltd"><%=XjDate(rs("sjjssj"),1)%></td>
  </tr>
  <tr>
    <td class="rtd">�鳤</td>
    <td colspan="3" class="ltd"><%If rs("zz")<>"" Then Response.Write(rs("zz")) else Response.Write(rs("jgzz")&"(�ṹ)��"&rs("sjzz")&"(���)")%></td>
    <td class="rtd">��������</td>
    <td class="ltd"><%=rs("jsdb")%></td>
  </tr>
</table>
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call mtask_userinfo(rs)%>
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call atask_userinfo(rs)%>
<%
End Function

Function mtask_userinfo(rs)
Dim strgy
strgy=""
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <%
			select case rs("mjxx") & rs("rwlr")
				case "ȫ�����"
				 If ((rs("gjfs")=3) and (rs("qhgj")=2)) or NullToNum(Rs("hgj"))<>0 Then%>
  <tr>
    <td class="ctd" width="10%">�󹲼��ṹ</td>
    <td class="ctd" width="*"><%=rs("gjjgr")%>&nbsp;</td>
    <td class="ctd" width="13%">�󹲼��ṹȷ��</td>
    <td class="ctd" width="*"><%=rs("gjjgshr")%>&nbsp;</td>
    <td class="ctd" width="10%">�󹲼����</td>
    <td class="ctd" width="*"><%=rs("gjsjr")%>&nbsp;</td>
    <%If rs("gjshr")<>"" Then%>
    <td class="ctd" width="10%">�󹲼����</td>
    <td class="ctd" width="*"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="*" colspan="4">&nbsp;</td>
    <%else%>
    <td class="ctd" width="13%">�󹲼�������</td>
    <td class="ctd" width="*" colspan="3"><%=rs("gjsjshr")%>&nbsp;</td>
    <%End If%>
  </tr>
  <% End If%>
  <%if (not isnull(rs("mtjgr"))) and (not isnull(rs("dxjgr"))) and rs("mtjgr")=rs("dxjgr") then%>
  <tr>
    <td class="ctd" width="10%" rowspan="2">ģ�߽ṹ</td>
    <%else%>
    <td class="ctd" width="10%">ģͷ�ṹ</td>
    <%end if%>
    <%if rs("mtjgr")=rs("dxjgr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtjgr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtjgr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtjgshr")=rs("dxjgshr") then%>
    <td class="ctd" width="11%" rowspan="2">�ṹȷ��</td>
    <%else%>
    <td class="ctd" width="11%">ģͷ�ṹȷ��</td>
    <%end if%>
    <%if rs("mtjgshr")=rs("dxjgshr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtjgshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtjgshr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="10%" rowspan="2">ģ�����</td>
    <%else%>
    <td class="ctd" width="10%">ģͷ���</td>
    <%end if%>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtsjr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtsjr")%>&nbsp;</td>
    <%end if%>
    <%If (not isnull(rs("mtshr"))) or (not isnull(rs("dxshr"))) Then%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="10%" rowspan="2">ģ�����</td>
    <%else%>
    <td class="ctd" width="10%">ģͷ���</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtshr")%>&nbsp;</td>
    <%end if%>
    <%else%>
    <%if rs("mtsjshr")=rs("dxsjshr") then%>
    <td class="ctd" width="11%" rowspan="2">������</td>
    <%else%>
    <td class="ctd" width="11%">ģͷ������</td>
    <%end if%>
    <%if rs("mtsjshr")=rs("dxsjshr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtsjshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtsjshr")%>&nbsp;</td>
    <%end if%>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="8%" rowspan="2">ģ��BOM</td>
    <%else%>
    <td class="ctd" width="10%">ģͷBOM</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="8%" rowspan="2"><%=rs("mtbomr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="8%"><%=rs("mtbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <tr>
    <%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
    <td class="ctd" width="10%">���ͽṹ</td>
    <%end if%>
    <%if isnull(rs("mtjgr")) or isnull(rs("dxjgr")) or (rs("mtjgr")<>rs("dxjgr")) then%>
    <td class="ctd" width="8%"><%=rs("dxjgr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtjgshr")) or isnull(rs("dxjgshr")) or (rs("mtjgshr")<>rs("dxjgshr")) then%>
    <td class="ctd" width="11%">���ͽṹȷ��</td>
    <%end if%>
    <%if isnull(rs("mtjgshr")) or isnull(rs("dxjgshr")) or (rs("mtjgshr")<>rs("dxjgshr")) then%>
    <td class="ctd" width="8%"><%=rs("dxjgshr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="10%">�������</td>
    <%end if%>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="8%"><%=rs("dxsjr")%>&nbsp;</td>
    <%end if%>
    <%If (not isnull(rs("mtshr"))) or (not isnull(rs("dxshr"))) Then%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="10%">�������</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="8%"><%=rs("dxshr")%>&nbsp;</td>
    <%end if%>
    <%else%>
    <%if isnull(rs("mtsjshr")) or isnull(rs("dxsjshr")) or (rs("mtsjshr")<>rs("dxsjshr")) then%>
    <td class="ctd" width="11%">����������</td>
    <%end if%>
    <%if isnull(rs("mtsjshr")) or isnull(rs("dxsjshr")) or (rs("mtsjshr")<>rs("dxsjshr")) then%>
    <td class="ctd" width="8%"><%=rs("dxsjshr")%>&nbsp;</td>
    <%end if%>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="10%">����BOM</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="8%"><%=rs("dxbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <%case "ȫ�׸���"	%>
  <%If Rs("gjsjr")<>"" Then%>
  <tr>
    <td class="ctd" width="10%" colspan="2">��</td>
    <td class="ctd" width="10%">��������</td>
    <td class="ctd" width="15%" colspan="2"><%=rs("gjsjr")%>&nbsp;</td>
    <td class="ctd" width="10%">�������</td>
    <td class="ctd" width="15%" colspan="2"><%=rs("gjshr")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="10%" rowspan="2">��</td>
    <td class="ctd" width="15%" rowspan="2">��</td>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="10%" rowspan="2">ģ�߸���</td>
    <%else%>
    <td class="ctd" width="10%">ģͷ����</td>
    <%end if%>
    <%if rs("mtsjr")=rs("dxsjr") then%>
    <td class="ctd" width="15%" rowspan="2"><%=rs("mtsjr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="15%"><%=rs("mtsjr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="10%" rowspan="2">ģ�����</td>
    <%else%>
    <td class="ctd" width="10%">ģͷ���</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="15%" rowspan="2"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="10%" rowspan="2">ģ��BOM</td>
    <%else%>
    <td class="ctd" width="10%">ģͷBOM</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="*" rowspan="2"><%=rs("mtbomr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <tr>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="10%">���͸���</td>
    <%end if%>
    <%if isnull(rs("mtsjr")) or isnull(rs("dxsjr")) or (rs("mtsjr")<>rs("dxsjr")) then%>
    <td class="ctd" width="15%"><%=rs("dxsjr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="10%">�������</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="10%">����BOM</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <%	case "ȫ�׸���"	%>
  <%If ((rs("gjfs")=3) and (rs("qhgj")=2)) or NullToNum(Rs("hgj"))<>0  Then%>
  <tr>
    <td class="ctd" width="50%">&nbsp;</td>
    <td class="ctd" width="10%">��������</td>
    <td class="ctd" width="15%" colspan="2"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width=* colspan="2">&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="50%">&nbsp;</td>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="10%" rowspan="2">ģ�߸���</td>
    <%else%>
    <td class="ctd" width="10%">ģͷ����</td>
    <%end if%>
    <%if rs("mtshr")=rs("dxshr") then%>
    <td class="ctd" width="15%" rowspan="2"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="10%" rowspan="2">ģ��BOM</td>
    <%else%>
    <td class="ctd" width="10%">ģͷBOM</td>
    <%end if%>
    <%if rs("mtbomr")=rs("dxbomr") then%>
    <td class="ctd" width="*" rowspan="2"><%=rs("mtbomr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%end if%>
  </tr>
  <tr>
    <td class="ctd" width="50%">&nbsp;</td>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="10%">���͸���</td>
    <%end if%>
    <%if isnull(rs("mtshr")) or isnull(rs("dxshr")) or (rs("mtshr")<>rs("dxshr")) then%>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="10%">����BOM</td>
    <%end if%>
    <%if isnull(rs("mtbomr")) or isnull(rs("dxbomr")) or (rs("mtbomr")<>rs("dxbomr")) then%>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%end if%>
    <%	case "ģͷ���"	%>
    <td class="ctd" width="10%">ģͷ�ṹ</td>
    <td class="ctd" width="10%"><%=rs("mtjgr")%>&nbsp;</td>
    <td class="ctd" width="11%">ģͷ�ṹȷ��</td>
    <td class="ctd" width="10%"><%=rs("mtjgshr")%>&nbsp;</td>
    <td class="ctd" width="10%">ģͷ���</td>
    <td class="ctd" width="10%"><%=rs("mtsjr")%>&nbsp;</td>
    <%If (not(isnull(Rs("mtshr")))) Then%>
    <td class="ctd" width="10%">ģͷ���</td>
    <td class="ctd" width="10%"><%=rs("mtshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="11%">ģͷ������</td>
    <td class="ctd" width="10%"><%=rs("mtsjshr")%>&nbsp;</td>
    <%End If%>
    <td class="ctd" width="10%">ģͷBOM</td>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%	case "ģͷ����"	%>
    <td class="ctd" width="10%">��</td>
    <td class="ctd" width="15%">��</td>
    <td class="ctd" width="10%">ģͷ����</td>
    <td class="ctd" width="15%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="10%">ģͷ���</td>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="10%">ģͷBOM</td>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%	case "ģͷ����"		%>
    <td class="ctd" width="10%">��</td>
    <td class="ctd" width="15%">��</td>
    <td class="ctd" width="10%">��</td>
    <td class="ctd" width="15%">��</td>
    <td class="ctd" width="10%">ģͷ����</td>
    <td class="ctd" width="15%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="10%">ģͷBOM</td>
    <td class="ctd" width="*"><%=rs("mtbomr")%>&nbsp;</td>
    <%	case "�������"		%>
    <td class="ctd" width="10%">���ͽṹ</td>
    <td class="ctd" width="10%"><%=rs("dxjgr")%>&nbsp;</td>
    <td class="ctd" width="11%">���ͽṹȷ��</td>
    <td class="ctd" width="10%"><%=rs("dxjgshr")%>&nbsp;</td>
    <td class="ctd" width="10%">�������</td>
    <td class="ctd" width="10%"><%=rs("dxsjr")%>&nbsp;</td>
    <%If (not(isnull(Rs("dxshr")))) Then%>
    <td class="ctd" width="10%">�������</td>
    <td class="ctd" width="10%"><%=rs("dxshr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="11%">����������</td>
    <td class="ctd" width="10%"><%=rs("dxsjshr")%>&nbsp;</td>
    <%End If%>
    <td class="ctd" width="10%">����BOM</td>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%	case "���͸���"		%>
    <td class="ctd" width="10%">��</td>
    <td class="ctd" width="15%">��</td>
    <td class="ctd" width="10%">���͸���</td>
    <td class="ctd" width="15%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="10%">�������</td>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="10%">����BOM</td>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%		case "���͸���"		%>
    <td class="ctd" width="10%">��</td>
    <td class="ctd" width="15%">��</td>
    <td class="ctd" width="10%">��</td>
    <td class="ctd" width="15%">��</td>
    <td class="ctd" width="10%">���͸���</td>
    <td class="ctd" width="15%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="10%">����BOM</td>
    <td class="ctd" width="*"><%=rs("dxbomr")%>&nbsp;</td>
    <%
				case else
				response.write(rs("mjxx") & rs("rwlr"))
			end select
		%>
  </tr>
</table>
<%
Response.Write(XjLine(5, "100%", ""))
If not(isNull(rs("gysjr"))) Then
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="35%" ><%=rs("gysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="35%" ><%=rs("gyshr")%>&nbsp;</td>
  </tr>
</table>
<%
else
select case rs("mjxx")
	case "ģͷ"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">ģͷ�������</td>
    <td class="ctd" width="35%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">ģͷ�������</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
  </tr>
</table>
<%
	case "����"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">���͹������</td>
    <td class="ctd" width="35%"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">���͹������</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
  </tr>
</table>
<%
	case else
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">ģͷ�������</td>
    <td class="ctd" width="35%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">ģͷ�������</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͹������</td>
    <td class="ctd" width="35%"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd" width="15%">���͹������</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
  </tr>
  <%If Rs("gjgysjr")<>"" Then%>
  <tr>
    <td class="ctd" width="10%">�����������</td>
    <td class="ctd" width="35%"><%=rs("gjgysjr")%>&nbsp;</td>
    <td class="ctd" width="10%">�����������</td>
    <td class="ctd"><%=rs("gjgyshr")%>&nbsp;</td>
  </tr>
  <%End If%>
</table>
<%
End select
End If
end function

function atask_userinfo(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <%
			select case rs("mjxx")
				case "ȫ��"
				%>
    <%if (not isnull(rs("mttsdr"))) and (not isnull(rs("dxtsdr"))) and (rs("mttsdr")=rs("dxtsdr")) then%>
    <td class="ctd" width="15%" rowspan="2">ģ�ߵ��Ե�</td>
    <%else%>
    <td class="ctd" width="15%">ģͷ���Ե�</td>
    <%end if%>
    <%if rs("mttsdr")=rs("dxtsdr") then%>
    <td class="ctd" width="10%" rowspan="2"><%=rs("mttsdr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="10%"><%=rs("mttsdr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mttsr")=rs("dxtsr") then%>
    <td class="ctd" width="15%" rowspan="2">ģ�ߵ���</td>
    <%else%>
    <td class="ctd" width="15%">ģͷ����</td>
    <%end if%>
    <%if rs("mttsr")=rs("dxtsr") then%>
    <td class="ctd" width="10%" rowspan="2"><%=rs("mttsr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="10%"><%=rs("mttsr")%>&nbsp;</td>
    <%end if%>
    <%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
    <td class="ctd" width="15%" rowspan="2">ģ�ߵ�����Ϣ����</td>
    <%else%>
    <td class="ctd" width="15%">ģͷ������Ϣ����</td>
    <%end if%>
    <%if rs("mttsxxzlr")=rs("dxtsxxzlr") then%>
    <td class="ctd" width="10%" rowspan="2"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <%else%>
    <td class="ctd" width="10%"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <%end if%>
    <td class="ctd" width="15%" rowspan="2">������Ϣ����</td>
    <td class="ctd" width="10%" rowspan="2"><%=rs("xtxxzlr")%>&nbsp;</td>
  </tr>
  <tr>
    <%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
    <td class="ctd" width="15%">���͵��Ե�</td>
    <%end if%>
    <%if isnull(rs("mttsdr")) or isnull(rs("dxtsdr")) or (rs("mttsdr")<>rs("dxtsdr")) then%>
    <td class="ctd" width="10%"><%=rs("dxtsdr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
    <td class="ctd" width="15%">���͵���</td>
    <%end if%>
    <%if isnull(rs("mttsr")) or isnull(rs("dxtsr")) or (rs("mttsr")<>rs("dxtsr")) then%>
    <td class="ctd" width="10%"><%=rs("dxtsr")%>&nbsp;</td>
    <%end if%>
    <%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
    <td class="ctd" width="15%">���͵�����Ϣ����</td>
    <%end if%>
    <%if isnull(rs("mttsxxzlr")) or isnull(rs("dxtsxxzlr")) or (rs("mttsxxzlr")<>rs("dxtsxxzlr")) then%>
    <td class="ctd" width="10%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <%end if%>
    <%
				case "ģͷ"
				%>
    <td class="ctd" width="15%">ģͷ���Ե�</td>
    <td class="ctd" width="15%"><%=rs("mttsdr")%>&nbsp;</td>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="15%"><%=rs("mttsr")%>&nbsp;</td>
    <td class="ctd" width="15%">ģͷ������Ϣ����</td>
    <td class="ctd" width="*"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="15%">������Ϣ����</td>
    <td class="ctd" width="*"><%=rs("xtxxzlr")%>&nbsp;</td>
    <%
				case "����"
				%>
    <td class="ctd" width="15%">���͵��Ե�</td>
    <td class="ctd" width="15%"><%=rs("dxtsdr")%>&nbsp;</td>
    <td class="ctd" width="15%">���͵���</td>
    <td class="ctd" width="15%"><%=rs("dxtsr")%>&nbsp;</td>
    <td class="ctd" width="15%">���͵�����Ϣ����</td>
    <td class="ctd" width="*"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="15%">������Ϣ����</td>
    <td class="ctd" width="*"><%=rs("xtxxzlr")%>&nbsp;</td>
    <%
				case else
				response.write(rs("mjxx"))
			end select
		%>
  </tr>
</table>
<%
end function


function mtask_alluserinfo(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <%select case rs("mjxx") & rs("rwlr")%>
  <%case "ȫ�����"%>
  <tr>
    <td class="ctd" width="15%">ģͷ�ṹ</td>
    <td class="ctd" width="9%"><%=rs("mtjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹ��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹ����</td>
    <td class="ctd" width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">ģͷ�ṹȷ��</td>
    <td class="ctd" width="9%"><%=rs("mtjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹȷ�Ͽ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹȷ�Ͻ���</td>
    <td class="ctd" width="20%"><%=rs("mtjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">ģͷ���</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��ƿ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��ƽ���</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">ģͷ���</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˽���</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">ģͷ������</td>
    <td class="ctd" width="9%"><%=rs("mtsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�����˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�����˽���</td>
    <td class="ctd" width="20%"><%=rs("mtsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">ģͷBOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM����</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���ͽṹ</td>
    <td class="ctd" width="9%"><%=rs("dxjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹ��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹ����</td>
    <td class="ctd" width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">���ͽṹȷ��</td>
    <td class="ctd" width="9%"><%=rs("dxjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹȷ�Ͽ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹȷ�Ͻ���</td>
    <td class="ctd" width="20%"><%=rs("dxjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">������ƿ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">������ƽ���</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">������˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">������˽���</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">����������</td>
    <td class="ctd" width="9%"><%=rs("dxsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">���������˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">���������˽���</td>
    <td class="ctd" width="20%"><%=rs("dxsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">����BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM����</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <% If (not isnull(rs("gjjgr"))) Then%>
  <tr>
    <td class="ctd" width="15%">�󹲼��ṹ</td>
    <td class="ctd" width="9%"><%=rs("gjjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">�󹲼��ṹ��ʼ</td>
    <td class="ctd" width="20%"><%=rs("gjjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">�󹲼��ṹ����</td>
    <td class="ctd" width="20%"><%=rs("gjjgjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">�󹲼����</td>
    <td class="ctd" width="9%"><%=rs("gjsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">�󹲼���ƿ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("gjsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">�󹲼���ƽ���</td>
    <td class="ctd" width="20%"><%=rs("gjsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">�󹲼����</td>
    <td class="ctd" width="9%"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">�󹲼���˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("gjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">�󹲼���˽���</td>
    <td class="ctd" width="20%"><%=rs("gjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <%case "ģͷ���"%>
  <tr>
    <td class="ctd" width="15%">ģͷ�ṹ</td>
    <td class="ctd" width="9%"><%=rs("mtjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹ��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹ����</td>
    <td class="ctd" width="20%"><%=rs("mtjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">ģͷ�ṹȷ��</td>
    <td class="ctd" width="9%"><%=rs("mtjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹȷ�Ͽ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�ṹȷ�Ͻ���</td>
    <td class="ctd" width="20%"><%=rs("mtjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">ģͷ���</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��ƿ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��ƽ���</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("mtshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">ģͷ���</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˽���</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">ģͷ������</td>
    <td class="ctd" width="9%"><%=rs("mtsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�����˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�����˽���</td>
    <td class="ctd" width="20%"><%=rs("mtsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">ģͷBOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM����</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <%case "�������"%>
  <tr>
    <td class="ctd" width="15%">���ͽṹ</td>
    <td class="ctd" width="9%"><%=rs("dxjgr")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹ��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxjgks")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹ����</td>
    <td class="ctd" width="20%"><%=rs("dxjgjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxjgshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">���ͽṹȷ��</td>
    <td class="ctd" width="9%"><%=rs("dxjgshr")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹȷ�Ͽ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxjgshks")%>&nbsp;</td>
    <td class="ctd" width="18%">���ͽṹȷ�Ͻ���</td>
    <td class="ctd" width="20%"><%=rs("dxjgshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">������ƿ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">������ƽ���</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <%If Rs("dxshr")<>"" Then%>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">������˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">������˽���</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <%else%>
  <tr>
    <td class="ctd" width="15%">����������</td>
    <td class="ctd" width="9%"><%=rs("dxsjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">���������˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxsjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">���������˽���</td>
    <td class="ctd" width="20%"><%=rs("dxsjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <tr>
    <td class="ctd" width="15%">����BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM����</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <%case "ȫ�׸���"%>
  <tr>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ŀ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ľ���</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷ���</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˽���</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷBOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM����</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͸���</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸��Ŀ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸��Ľ���</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">������˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">������˽���</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">����BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM����</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <% If (not isnull(rs("gjsjr"))) Then%>
  <tr>
    <td class="ctd" width="15%">��������</td>
    <td class="ctd" width="9%"><%=rs("gjsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">�������Ŀ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("gjsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">�������Ľ���</td>
    <td class="ctd" width="20%"><%=rs("gjsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">������˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("gjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">������˽���</td>
    <td class="ctd" width="20%"><%=rs("gjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <%case "ģͷ����"%>
  <tr>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ŀ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ľ���</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="9%"><%=rs("mtsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ŀ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ľ���</td>
    <td class="ctd" width="20%"><%=rs("mtsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷ���</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ��˽���</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷBOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM����</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <%case "���͸���"%>
  <tr>
    <td class="ctd" width="15%">���͸���</td>
    <td class="ctd" width="9%"><%=rs("dxsjr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸��Ŀ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxsjks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸��Ľ���</td>
    <td class="ctd" width="20%"><%=rs("dxsjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">������˿�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">������˽���</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">����BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM����</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <%case "ȫ�׸���"%>
  <tr>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���鿪ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�������</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷBOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM����</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͸���</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸��鿪ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸������</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">����BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM����</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <% If (not isnull(rs("gjshr"))) Then%>
  <tr>
    <td class="ctd" width="15%">��������</td>
    <td class="ctd" width="9%"><%=rs("gjshr")%>&nbsp;</td>
    <td class="ctd" width="18%">�������鿪ʼ</td>
    <td class="ctd" width="20%"><%=rs("gjshks")%>&nbsp;</td>
    <td class="ctd" width="18%">�����������</td>
    <td class="ctd" width="20%"><%=rs("gjshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
  <%case "ģͷ����"%>
  <tr>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="9%"><%=rs("mtshr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���鿪ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtshks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ�������</td>
    <td class="ctd" width="20%"><%=rs("mtshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷBOM</td>
    <td class="ctd" width="9%"><%=rs("mtbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("mtbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷBOM����</td>
    <td class="ctd" width="20%"><%=rs("mtbomjs")%>&nbsp;</td>
  </tr>
  <%case "���͸���"%>
  <tr>
    <td class="ctd" width="15%">���͸���</td>
    <td class="ctd" width="9%"><%=rs("dxshr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸��鿪ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxshks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͸������</td>
    <td class="ctd" width="20%"><%=rs("dxshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">����BOM</td>
    <td class="ctd" width="9%"><%=rs("dxbomr")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM��ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxbomks")%>&nbsp;</td>
    <td class="ctd" width="18%">����BOM����</td>
    <td class="ctd" width="20%"><%=rs("dxbomjs")%>&nbsp;</td>
  </tr>
  <%end select%>
</table>
<%dim strgy
strgy=""
Response.Write(XjLine(5, "100%", ""))
If not(isNull(rs("gysjr"))) Then
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%" ><%=rs("gysjr")%>&nbsp;</td>
    <td class="ctd" width="18%">��ʼʱ��</td>
    <td class="ctd" width="20%"><%=rs("gysjks")%>&nbsp;</td>
    <td class="ctd" width="18%">����ʱ��</td>
    <td class="ctd" width="20%"><%=rs("gysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">�������</td>
    <td class="ctd" width="9%" ><%=rs("gyshr")%>&nbsp;</td>
    <td class="ctd" width="18%">��ʼʱ��</td>
    <td class="ctd" width="20%"><%=rs("gyshks")%>&nbsp;</td>
    <td class="ctd" width="18%">����ʱ��</td>
    <td class="ctd" width="20%"><%=rs("gyshjs")%>&nbsp;</td>
  </tr>
</table>
<%
else
select case rs("mjxx")
	case "ģͷ"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">ģͷ�������</td>
    <td class="ctd" width="9%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" >��ʼʱ��</td>
    <td class="ctd" width="20%"><%=rs("mtgysjks")%>&nbsp;</td>
    <td class="ctd" >����ʱ��</td>
    <td class="ctd" width="20%"><%=rs("mtgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">ģͷ�������</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
    <td class="ctd" >����ʱ��</td>
    <td class="ctd"><%=rs("mtgyshks")%>&nbsp;</td>
    <td class="ctd" >����ʱ��</td>
    <td class="ctd"><%=rs("mtgyshjs")%>&nbsp;</td>
  </tr>
</table>
<%
	case "����"
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd">���͹������</td>
    <td class="ctd"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgysjks")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">���͹������</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgyshks")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgyshjs")%>&nbsp;</td>
  </tr>
</table>
<%
	case else
	%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <tr>
    <td class="ctd" width="15%">ģͷ�������</td>
    <td class="ctd" width="9%"><%=rs("mtgysjr")%>&nbsp;</td>
    <td class="ctd" >��ʼʱ��</td>
    <td class="ctd" width="20%"><%=rs("mtgysjks")%>&nbsp;</td>
    <td class="ctd" >����ʱ��</td>
    <td class="ctd" width="20%"><%=rs("mtgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">ģͷ�������</td>
    <td class="ctd"><%=rs("mtgyshr")%>&nbsp;</td>
    <td class="ctd" >����ʱ��</td>
    <td class="ctd"><%=rs("mtgyshks")%>&nbsp;</td>
    <td class="ctd" >����ʱ��</td>
    <td class="ctd"><%=rs("mtgyshjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">���͹������</td>
    <td class="ctd"><%=rs("dxgysjr")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgysjks")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">���͹������</td>
    <td class="ctd"><%=rs("dxgyshr")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgyshks")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("dxgyshjs")%>&nbsp;</td>
  </tr>
  <%If Rs("gjgysjr")<>"" Then%>
  <tr>
    <td class="ctd">�����������</td>
    <td class="ctd"><%=rs("gjgysjr")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("gjgysjks")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("gjgysjjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd">�����������</td>
    <td class="ctd"><%=rs("gjgyshr")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("gjgyshks")%>&nbsp;</td>
    <td class="ctd">����ʱ��</td>
    <td class="ctd"><%=rs("gjgyshjs")%>&nbsp;</td>
  </tr>
  <%End If%>
</table>
<%
End select
End If
end function

function atask_alluserinfo(rs)
%>
<table class="xtable" cellspacing="0" cellpadding="3" width="95%">
  <%select case rs("mjxx")%>
  <%case "ȫ��"%>
  <tr>
    <td class="ctd" width="15%">ģͷ���Ե�</td>
    <td class="ctd" width="9%"><%=rs("mttsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ե���ʼ</td>
    <td class="ctd" width="20%"><%=rs("mttsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ե�����</td>
    <td class="ctd" width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="9%"><%=rs("mttsr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Կ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mttsks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Խ���</td>
    <td class="ctd" width="20%"><%=rs("mttsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷ������Ϣ����</td>
    <td class="ctd" width="9%"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ������Ϣ����ʼ</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ������Ϣ�������</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͵��Ե�</td>
    <td class="ctd" width="9%"><%=rs("dxtsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Ե���ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Ե�����</td>
    <td class="ctd" width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͵���</td>
    <td class="ctd" width="9%"><%=rs("dxtsr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Կ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxtsks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Խ���</td>
    <td class="ctd" width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͵�����Ϣ����</td>
    <td class="ctd" width="9%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵�����Ϣ����ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵�����Ϣ�������</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">������Ϣ����</td>
    <td class="ctd" width="9%"><%=rs("xtxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">������Ϣ����ʼ</td>
    <td class="ctd" width="20%"><%=rs("xtxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">������Ϣ�������</td>
    <td class="ctd" width="20%"><%=rs("xtxxsjjs")%>&nbsp;</td>
  </tr>
  <%case "ģͷ"%>
  <tr>
    <td class="ctd" width="15%">ģͷ���Ե�</td>
    <td class="ctd" width="9%"><%=rs("mttsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ե���ʼ</td>
    <td class="ctd" width="20%"><%=rs("mttsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Ե�����</td>
    <td class="ctd" width="20%"><%=rs("mttsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷ����</td>
    <td class="ctd" width="9%"><%=rs("mttsr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Կ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("mttsks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ���Խ���</td>
    <td class="ctd" width="20%"><%=rs("mttsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">ģͷ������Ϣ����</td>
    <td class="ctd" width="9%"><%=rs("mttsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ������Ϣ����ʼ</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">ģͷ������Ϣ�������</td>
    <td class="ctd" width="20%"><%=rs("mttsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">������Ϣ����</td>
    <td class="ctd" width="9%"><%=rs("xtxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">������Ϣ����ʼ</td>
    <td class="ctd" width="20%"><%=rs("xtxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">������Ϣ�������</td>
    <td class="ctd" width="20%"><%=rs("xtxxsjjs")%>&nbsp;</td>
  </tr>
  <%case "����"%>
  <tr>
    <td class="ctd" width="15%">���͵��Ե�</td>
    <td class="ctd" width="9%"><%=rs("dxtsdr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Ե���ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxtsdks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Ե�����</td>
    <td class="ctd" width="20%"><%=rs("dxtsdjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͵���</td>
    <td class="ctd" width="9%"><%=rs("dxtsr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Կ�ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxtsks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵��Խ���</td>
    <td class="ctd" width="20%"><%=rs("dxtsjs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">���͵�����Ϣ����</td>
    <td class="ctd" width="9%"><%=rs("dxtsxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵�����Ϣ����ʼ</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">���͵�����Ϣ�������</td>
    <td class="ctd" width="20%"><%=rs("dxtsxxzljs")%>&nbsp;</td>
  </tr>
  <tr>
    <td class="ctd" width="15%">������Ϣ����</td>
    <td class="ctd" width="9%"><%=rs("xtxxzlr")%>&nbsp;</td>
    <td class="ctd" width="18%">������Ϣ����ʼ</td>
    <td class="ctd" width="20%"><%=rs("xtxxzlks")%>&nbsp;</td>
    <td class="ctd" width="18%">������Ϣ�������</td>
    <td class="ctd" width="20%"><%=rs("xtxxsjjs")%>&nbsp;</td>
  </tr>
  <%
				case else
				response.write(rs("mjxx"))
			end select
			%>
  <tr>
    <td>
  </tr>
    </td>

    </tr>

</table>
<%
end function

Function DisTd(strfield1,strfield2,strdiff,rs)
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <%if isnull(strfield2) then
  	  if datediff("d", now, rs("jhjssj")) < strdiff then%>
    <td height="7" bgcolor="red"></td>
    <%else
   		 	if isnull(strfield1) then%>
    <td height="7"></td>
    <%else%>
    <td height="7" bgcolor="#6F87CC"></td>
    <%end if
   	 end if
    else
    	if datediff("d", strfield2, rs("jhjssj")) < strdiff then%>
    <td height="7" bgcolor="#991100"></td>
    <%else%>
    <td height="7" bgcolor="#338833"></td>
    <%end if
    end if%>
  </tr>
</table>
<%
End Function

'�ṹ,��ƶ�����ʱ16:12 2007-4-1-������
Function DisTdjg(strfield1,strfield2,strdiff,rs)
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <%if isnull(strfield2) then
  		  if datediff("d", strdiff, now) > 0 then%>
    <td height="7" bgcolor="red"></td>
    <%else
   		 	if isnull(strfield1) then%>
    <td height="7">&nbsp;</td>
    <%else%>
    <td height="7" bgcolor="#6F87CC"></td>
    <%end if
   		 end if
 	 else
		if datediff("d", strdiff, strfield2) > 0 then%>
    <td height="7" bgcolor="#991100"></td>
    <%else%>
    <td height="7" bgcolor="#338833"></td>
    <%end if
  	end if%>
  </tr>
</table>
<%
End Function

Function DisTd2(strfield1,strfield2,rs)
%>
<table width="100%" cellpadding="0" cellspacing="0">
  <tr>
    <%if isnull(strfield2) then
    	if isnull(strfield1) then%>
    <td height="7"></td>
    <%else%>
    <td height="7" bgcolor="#6F87CC"></td>
    <%end if
    else%>
    <td height="7" bgcolor="#338833"></td>
    <%end if%>
  </tr>
</table>
<%
End Function

Function CutLine()	'ͼ��
%>
<table width="95%" cellpadding="0" cellspacing="0" border="0">
  <tr>
    <td align="right"><table cellpadding="2" cellspacing="2" border="0">
        <tr>
          <td align="right">ͼ��:</td>
          <td class="ctd" bgcolor="#6F87CC"><font color="white">�ƻ���ִ��</font></td>
          <td class="ctd" bgcolor="#338833"><font color="white">�������</font></td>
          <td class="ctd" bgcolor="#991100"><font color="white">�������</font></td>
          <td class="ctd" bgcolor="#ff0000"><font color="white">����δ���</font></td>
        </tr>
      </table></td>
  </tr>
</table>
<%
End Function

Function DisFzInfo(Rs)	'��ʾ��ֵ��Ϣ
	Dim Dismtfz, Disdxfz, Disgjfz, DisTslb, DisTsxx, DisTssx, DisSql, mtgjf, dxgjf, DisRs, ssgjf, qbfgjf, qgjf, hgjf
	Dismtfz="" : Disdxfz="" : Disgjfz="" : DisTslb="" : DisTsxx=0 : DisTssx=0 : mtgjf=0 : dxgjf=0
	ssgjf=NullToNum(Rs("ssgj"))
	qbfgjf=NullToNum(Rs("qbfgj"))
	qgjf=NullToNum(Rs("qgj"))
	hgjf=NullToNum(Rs("hgj"))

	select case ssgjf&qbfgjf&qgjf&hgjf
		Case "0000"			'����08�湲���Ʒ�ģʽ
			'ֻ����Ӳǰ�����ķ�ֵ�Ų��ּӵ�ģͷ���ּӵ�������
			if Rs("gjfs")="3" and Rs("qhgj")="1" Then
				Dismtfz=Rs("mjzf")*Rs("mtbl")/100
				Disdxfz=Rs("mjzf")*(100-Rs("mtbl"))/100
			End if
			'��Ӳ�󹲼��ķ�ֵ�����ӵ��󹲼�����
			If Rs("gjfs")="3" and Rs("qhgj")="2" Then
				Dismtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100
				Disdxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
				Disgjfz=Rs("gjzf")
			End if
			'�������������й������ȫ�ӵ�ģͷ
			If (not (Rs("gjfs")="3")) Then
				Dismtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + Rs("gjzf")
				Disdxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100
			End if
		Case Else		'09�湲���Ʒ�ģʽ
			If qgjf<>0 Then
				mtgjf=qgjf*Rs("mtbl")/100
				dxgjf=qgjf-mtgjf
			End If
			mtgjf=mtgjf+ssgjf+qbfgjf
			Dismtfz=(Rs("mjzf")-Rs("gjzf"))*Rs("mtbl")/100 + mtgjf
			Disdxfz=(Rs("mjzf")-Rs("gjzf"))*(100-Rs("mtbl"))/100 + dxgjf
			Disgjfz=Rs("hgj")
	end select
	'�������
	DisTslb=Rs("TSLB")
	If ((rs("gjfs")=3) and (rs("qhgj")=2)) or NullToNum(Rs("hgj"))<>0  Then
		DisFzInfo="ģ���ܷ�: <b>" & Rs("mjzf") & "</b> ��<br>" &_
		"ģͷ��ֵ: <b>" & Dismtfz & "</b> ��<br>" &_
		"���ͷ�ֵ: <b>" & Disdxfz & "</b> ��<br>" &_
		"�󹲼���ֵ: <b>" & Disgjfz & "</b> ��<br>" &_
		"BOM��ֵ: <b>" & Rs("bomzf") & "</b> ��<br>" &_
		"���Ե���ֵ: <b>" & Rs("tsdzf") & "</b> ��<br>" &_
		"���Է�ֵ: <b>" & Rs("tszf") & "</b> ��<br>" &_
		"������Ϣ�����ֵ: <b>" & Rs("tsxxzlzf") & "</b> ��"
	Else
	DisFzInfo="ģ���ܷ�: <b>" & Rs("mjzf") & "</b> ��<br>" &_
		"ģͷ��ֵ: <b>" & Dismtfz & "</b> ��<br>" &_
		"���ͷ�ֵ: <b>" & Disdxfz & "</b> ��<br>" &_
		"BOM��ֵ: <b>" & Rs("bomzf") & "</b> ��<br>" &_
		"���Ե���ֵ: <b>" & Rs("tsdzf") & "</b> ��<br>" &_
		"���Է�ֵ: <b>" & Rs("tszf") & "</b> ��<br>" &_
		"������Ϣ�����ֵ: <b>" & Rs("tsxxzlzf") & "</b> ��"
	End if
	If not(isNull(DisTslb)) Then
		DisSql="select * from [c_tscs] where dmlb like '%"&DisTslb&"%'"
		Set DisRs=xjweb.Exec(DisSql, 1)
		If not DisRs.eof Then
			DisTsxx=DisRs("edxx")
			DisTssx=DisRs("edsx")
		End If
		DisRs.Close
		set DisRs = nothing
		set DisSql = nothing
		DisFzInfo=DisFzInfo &"<br>����Դ���: <b>"& DisTsxx &" - "& DisTssx &"</b>��"
	End If
End Function
%>
