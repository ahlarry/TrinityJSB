<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'15:28 2006-11-2-������
Call ChkPageAble(0)
CurPage="������� �� ��ӡ������"
strPage="mtask"
xjweb.header()
'Call TopTable()
Call Main()
'Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<Table border=0 cellspacing=0 cellpadding=2 width="720">
  <Tr>
    <Td height=300><%Call mtaskDisplay()%>
      <%Response.Write(XjLine(10,"100%",""))%></Td>
  </Tr>
</Table>
<%
End Sub

Function mtaskDisplay()
	Dim s_lsh, action
	s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("��ȷ����ӡ���������ˮ��!") : Exit Function
	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then Call JsAlert("��ˮ�� ��" & s_lsh & "�� �����鲻����! ������������ˮ��!", "mtask_display.asp") : Exit Function
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
<%Call TbTopic("����ģ�߳�����ģ���������")%>
<table class=ktable cellspacing=0 cellpadding=3 width="95%">
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8"><b>��ͬ��Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd" width="13%">������</td>
    <td class="ltd"><%=rs("ddh")%></td>
    <td class="rtd" width="13%">��ˮ��</td>
    <td colspan="2" class="ltd" width="*"><a href="SreacDwg.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    <td class="rtd" width="13%">ģ��</td>
    <td colspan="2" class="ltd" width="*"><%=rs("mh")%></td>
  </tr>
  <tr>
    <td class="rtd">�ͻ�����</td>
    <td class="ltd"><%=rs("dwmc")%></td>
    <td class="rtd">��������</td>
    <td colspan="2" class="ltd"><%=rs("dmmc")%></td>
    <td class="rtd">ģ�߲���</td>
    <td colspan="2" class="ltd"><%=rs("mjcl")%></td>
  </tr>
  <tr>
    <td class="rtd">�豸����</td>
    <td class="ltd"><%=rs("sbcj")%></td>
    <td class="rtd">ˮ��ͷ����</td>
    <td colspan="2" class="ltd"><%=rs("sjtsl")%></td>
    <td class="rtd">����ͷ����</td>
    <td colspan="2" class="ltd"><%=rs("qjtsl")%></td>
  </tr>
  <tr>
    <td class="rtd">�������ͺ�</td>
    <td class="ltd"><%=rs("jcjxh")%></td>
    <td class="rtd">��������</td>
    <td colspan="2" class="ltd"><%=rs("jcfx")%></td>
    <td class="rtd">ǣ���ٶ�</td>
    <td colspan="2" class="ltd"><%=rs("qysd")%> ��/��(m/min)</td>
  </tr>
  <tr>
    <td class="rtd">����Ȱ�</td>
    <td class="ltd"><%if rs("pjrb") then%>
      ��
      <%else%>
      ��
      <%end if%></td>
    <td class="rtd">���Ȱ���Ϣ</td>
    <td colspan="2" class="ltd">����:<%=rs("jrbxs")%> ����:<%=rs("jrbcl")%> &nbsp;&nbsp;<%=rs("jrbxx")%></td>
    <td class="rtd">ǻ��</td>
    <td colspan="2" class="ltd"><%=rs("qs")%>ǻ</td>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8"><b>ģ����Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd"  width="13%">��������</td>
    <td class="ltd"><%=rs("mjxx") & rs("rwlr")%></td>
        <td class="rtd"  width="13%">���ڵ���</td>
    <td class="ltd"  colspan="2" ><%if rs("cnts") then%>
      ��
      <%else%>
      &nbsp;/
      <%end if%></td>
    <td class="rtd"  width="13%">�������</td>
    <% If Rs("cnts") Then%>
    <%If Not(isnull(Rs("tslb"))) Then%>
    <td class="ltd"  colspan="2"><a href="mtest_display.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("tslb")%></a></td>
    <%Else%>
    <td class="ltd" colspan="2" >&nbsp;/</td>
    <%End If%>
    <%Else%>
    <%If Rs("beit") Then%>
    <td class="ltd"  colspan="2">����</td>
    <%Else%>
    <td class="ltd" colspan="2">&nbsp;/</td>
    <%End If%>
    <%End If%>
  </tr>
  <tr>
    <td class="rtd">ģͷ�ṹ</td>
    <td class="ltd"><%if IsNull(rs("mtjg")) Then
    	Response.Write("&nbsp;")
    else
    	Response.Write(rs("mtjg"))
    End if%></td>
    <td class="rtd">���ͽṹ</td>
    <td class="ltd" colspan="2" ><%if IsNull(rs("dxjg")) Then
    	Response.Write("&nbsp;/")
    else
    	Response.Write(rs("dxjg"))
    End if%></td>
    <td class="rtd">ˮ��ṹ</td>
    <td class="ltd" colspan="2" ><%if IsNull(rs("sxjg")) Then
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
    <td class="ltd" colspan="2" ><%=rs("mtljcc")%></td>
    <td class="rtd">�ȵ�ż���</td>
    <td class="ltd" colspan="2" ><%=rs("rdogg")%></td>
  </tr>
  <tr>
    <td class="rtd">��������</td>
    <td class="ltd"><%=Trim(gjxx)%></td>
    <td class="rtd">�������ӳߴ�</td>
    <td class="ltd" colspan="2" ><%=rs("gjljcc")%>&nbsp;</td>
    <td class="rtd">�Ͳıں�</td>
    <td class="ltd" colspan="2" ><%=Rs("xcbh")%>����</td>
  </tr>
  <tr>
  </tr>
  <tr bgcolor="#DDDDDD">
    <td class="ltd" height="25" colspan="8" alt="<%=DisFzInfo(Rs)%>"><b>������Ϣ</b></td>
  </tr>
  <tr>
    <td class="rtd" height="180">�����¼</td>
    <td class="ltd"colspan="7"><table width="100%" >
        <tr>
          <td><label>
              <input type="checkbox" name="psyx" value="3" id="psyx_2" />
              �����󣬽������:</label></td>
          <td><label>
              <input type="checkbox" name="psyx" value="1" id="psyx_0" />
              ������ƹ淶�������</label></td>
          <td><label>
              <input type="checkbox" name="psyx" value="2" id="psyx_1" />
              ���տͻ���ͼ������� </label></td>
        </tr>
        <tr>
          <td height="120" colspan="5" valign="middle"><%=xjweb.HtmlToCode(rs("psjl"))%></Td>
        </tr>
        <tr valign="bottom">
          <td></Td>
          <td colspan="4" valign="bottom"><p align="left">ǩ��:
            </p></Td>
        </tr>
      </table></Td>
  </tr>
  <tr>
    <td class="rtd">��ע</td>
    <td class="ltd" colspan="7" height="180" valign="top"><%=xjweb.HtmlToCode(rs("bz"))%></td>
  </tr>
  <tr>
    <td class="rtd">�ƻ���ʼ</td>
    <%If rs("jhkssj")<>"" Then%>
    <td class="ltd"><%=XjDate(rs("jhkssj"),3)%></td>
    <%else%>
    <td class="ltd" width="120">&nbsp;/&nbsp;</td>
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
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call mtask_userinfo(rs)%>
<%Response.Write(XjLine(5,web_info(8),""))%>
<%Call atask_userinfo(rs)%>
<%
	Rs.Close
End Function
%>
