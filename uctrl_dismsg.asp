<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="�û����� �� �鿴ϵͳ����Ϣ"
strPage="uctrl"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()

dim action, strbox
action="" : strbox=""
action=request("action")
if trim(request("box"))<>"" then strbox=trim(request("box"))

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>

<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
  <Tr>
    <Td class=ctd height=300><%Call uctrlDismsg()%>
      <%Response.Write(XjLine(10,"100%",""))%>
    </Td>
  </Tr>
</Table>
<%
End Sub

Function uctrlDismsg()
	select case strbox
		case "send"
			call dis_sendbox()
		case "incept"
			call dis_inceptbox()
		case else
			call dis_sendbox()
	end select
end function

function dis_inceptbox()
	dim bhavemsg, strfsr		'�Ƿ��ж���
	bhavemsg=false
	strfsr=Request("fsr")
	response.write(tbtopic("�ռ���"))
%>
<Table border=0 cellpadding=2 cellspacing=0 width="100%" align="center">
  <Form action="<%=Request.Servervariables("script_name")%>" method=get>
    <Tr>
      <Td>ɸѡ����&nbsp;����&nbsp;&nbsp;������:
        <Select name="fsr" onchange='window.location.href("<%=request.servervariables("script_name")%>?box=incept&fsr=" + this.value);'>
          <option value="">ȫ��</option>
          <%for i = 0 to ubound(c_alluser)%>
          <option value='<%=c_alluser(i)%>' <%If strfsr=c_alluser(i) Then%> Selected<%End If%>><%=c_alluser(i)%></option>
          <%next%>
        </Select>
      </Td>
    </Tr>
  </Form>
</Table>
<table border="0" cellspacing="0" cellpadding="3" class=xtable width="98%" align="center">
  <tr>
    <td class=ltd colspan=6><b> �� �ռ���</b></td>
  </tr>
  <tr>
    <th class=th width="30">״̬</th>
    <th class=th width="100">������</th>
    <th class=th width="300">����</th>
    <th class=th width="*">ʱ��</th>
    <th class=th width="80">��С</th>
    <th class=th width="30">����</th>
  </tr>
  <form name="msgbox" action="msg_indb.asp?action=delete" method="post" onsubmit="return checkmsgbox();">

  <%If strfsr="" Then
			strSql="select * from ims_message where incept='"&Session("userName")&"' and delr<>2 order by id desc"
		else
			strSql="select * from ims_message where sender='"&strfsr&"' and incept='"&Session("userName")&"' and delr<>2 order by id desc"
		End If
			set rs=xjweb.Exec(strSql, 1)
			do while not rs.eof
		%>
  <tr <%if rs("delr")=0 then%>bgcolor="#dddddd"<%end if%>>
    <td class=ctd><%if rs("delr")=0 then%>
      <img src="images/uctrl/m_news.gif" border="0">
      <%else%>
      <img src="images/uctrl/m_olds.gif" border="0">
      <%end if%></td>
    <td class=ctd><%=rs("sender")%></td>
    <td class=ltd><a href="msg_dis.asp?<%if rs("delr")=0 then%>new=true&<%end if%>id=<%=rs("id")%>" target="_blank"><%=rs("title")%></a></td>
    <td class=ctd><%=rs("sendtime")%></td>
    <td class=ctd ><%=len(rs("content"))%>Byte</td>
    <td class=ctd ><input type=checkbox class="radio" id=chkbox name=chkbox value=<%=rs("id")%>></td>
  </tr>
  <%
				bhavemsg=true
				rs.movenext
			loop
			rs.close

			if bhavemsg then
		%>
  <tr>
    <td class=rtd colspan="6"> ͼ��:<img src="images/uctrl/m_news.gif" border="0">δ�� <img src="images/uctrl/m_olds.gif" border="0">�Ѷ�&nbsp;&nbsp;&nbsp;&nbsp; ��ʡ�������ռ���һ������! �뼰ʱɾ����ʱ��Ϣ! &nbsp;&nbsp;
      <input type="checkbox" id="selectall" name="selallmsg" onclick="chkselall(this.form);">
      <label for="selectall">ȫѡ</label>
      &nbsp;&nbsp;
      <input type="submit" value=" ɾ��ѡ�� "></td>
  </tr>
  <%else%>
  <tr>
    <td class=rtd colspan="6" height=30><b>�ռ�����ʱû�ж���</b></td>
  </tr>
  <%end if%>
  <input type="hidden" name="boxkind" value="incept">
  <form>

</table>
<%
end function

function dis_sendbox()
	dim bhavemsg		'�Ƿ��ж���
	bhavemsg=false
	response.write(tbtopic("������"))
%>
<table border="0" cellspacing="0" cellpadding="3" class=xtable width="98%" align="center">
  <tr>
    <td class=ltd colspan=6><b> �� ������</b></td>
  </tr>
  <tr>
    <th class=th width="30">״̬</th>
    <th class=th width="100">������</th>
    <th class=th width="300">����</th>
    <th class=th width="*">ʱ��</th>
    <th class=th width="80">��С</th>
    <th class=th width="30">����</th>
  </tr>
  <form name="msgbox" action="msg_indb.asp?action=delete" method="post" onsubmit="return checkmsgbox();">

  <%
			strSql="select * from [ims_message] where sender='"&Session("userName")&"' and dels<>2 order by id desc"
			set rs=xjweb.Exec(strSql, 1)
			do while not rs.eof
		%>
  <tr <%if rs("delr")=0 then%>bgcolor="#dddddd"<%end if%>>
    <td class=ctd><%if rs("delr")=0 then%>
      <img src="images/uctrl/m_news.gif" border="0">
      <%else%>
      <img src="images/uctrl/m_olds.gif" border="0">
      <%end if%></td>
    <td class=ctd><%=rs("incept")%></td>
    <td class=ltd ><a href="msg_dis.asp?action=send&id=<%=rs("id")%>" target="_blank"><%=rs("title")%></a></td>
    <td class=ctd><%=rs("sendtime")%></td>
    <td class=ctd ><%=len(rs("content"))%>Byte</td>
    <td class=ctd ><input type=checkbox class="radio" id=chkbox name=chkbox value=<%=rs("id")%>></td>
  </tr>
  <%
				bhavemsg=true
				rs.movenext
			loop
			rs.close

			if bhavemsg then
		%>
  <tr>
    <td class=rtd colspan="6"> ͼ��:<img src="images/uctrl/m_news.gif" border="0">�Է�δ�� <img src="images/uctrl/m_olds.gif" border="0">�Է��Ѷ� &nbsp;&nbsp;&nbsp;&nbsp; ��ʡ�������ռ���һ������! �뼰ʱɾ����ʱ��Ϣ!  &nbsp;&nbsp;
      <input type="checkbox" id="selectall" name="selallmsg" onclick="chkselall(this.form);">
      <label for="selectall">ȫѡ</label>
      &nbsp;&nbsp;
      <input type="submit" value=" ɾ��ѡ�� "></td>
  </tr>
  <%else%>
  <tr>
    <td class=ctd colspan="6" height=30><b>��������ʱû�ж���</b></td>
  </tr>
  <%end if%>
  <input type="hidden" name="boxkind" value="send">
  <form>

</table>
<%
end function
%>
<script language="javascript">
	function checkmsgbox()
	{
		var total = 0;
		var str="";
		var max = document.all.chkbox.length;
		for (var idx = 0; idx < max; idx++)
		{
			if (eval("document.all.chkbox[" + idx + "].checked") == true)
			{
				total += 1;
				str+=eval("document.all.chkbox[" + idx + "].value")
			}
		}
		if(typeof(max)=="undefined" && document.all.chkbox.checked==true) total=1;
		if(total>0)
			return confirm("��׼��ɾ��ѡ��� " + total + " ����Ϣ! ɾ���󽫲��ָܻ�,ȷ����?");
		else
			{alert("����ѡ����Ҫɾ������Ϣ!");return false;}
	}

	function chkselall(form)
	{
		 for (var i=0;i<form.elements.length;i++)
		 {
			var e = form.elements[i];
			if (e.name != 'selallmsgl')
				e.checked = form.selallmsg.checked;
		}
	}
</script>
