<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
CurPage="用户操作 → 查看系统短消息"
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
	dim bhavemsg, strfsr		'是否有短信
	bhavemsg=false
	strfsr=Request("fsr")
	response.write(tbtopic("收件箱"))
%>
<Table border=0 cellpadding=2 cellspacing=0 width="100%" align="center">
  <Form action="<%=Request.Servervariables("script_name")%>" method=get>
    <Tr>
      <Td>筛选条件&nbsp;——&nbsp;&nbsp;发送人:
        <Select name="fsr" onchange='window.location.href("<%=request.servervariables("script_name")%>?box=incept&fsr=" + this.value);'>
          <option value="">全部</option>
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
    <td class=ltd colspan=6><b> ■ 收件箱</b></td>
  </tr>
  <tr>
    <th class=th width="30">状态</th>
    <th class=th width="100">发件人</th>
    <th class=th width="300">主题</th>
    <th class=th width="*">时间</th>
    <th class=th width="80">大小</th>
    <th class=th width="30">操作</th>
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
    <td class=rtd colspan="6"> 图例:<img src="images/uctrl/m_news.gif" border="0">未读 <img src="images/uctrl/m_olds.gif" border="0">已读&nbsp;&nbsp;&nbsp;&nbsp; 节省服务器空间是一种美德! 请及时删除过时信息! &nbsp;&nbsp;
      <input type="checkbox" id="selectall" name="selallmsg" onclick="chkselall(this.form);">
      <label for="selectall">全选</label>
      &nbsp;&nbsp;
      <input type="submit" value=" 删除选择 "></td>
  </tr>
  <%else%>
  <tr>
    <td class=rtd colspan="6" height=30><b>收件箱暂时没有短信</b></td>
  </tr>
  <%end if%>
  <input type="hidden" name="boxkind" value="incept">
  <form>

</table>
<%
end function

function dis_sendbox()
	dim bhavemsg		'是否有短信
	bhavemsg=false
	response.write(tbtopic("发件箱"))
%>
<table border="0" cellspacing="0" cellpadding="3" class=xtable width="98%" align="center">
  <tr>
    <td class=ltd colspan=6><b> ■ 发件箱</b></td>
  </tr>
  <tr>
    <th class=th width="30">状态</th>
    <th class=th width="100">发件人</th>
    <th class=th width="300">主题</th>
    <th class=th width="*">时间</th>
    <th class=th width="80">大小</th>
    <th class=th width="30">操作</th>
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
    <td class=rtd colspan="6"> 图例:<img src="images/uctrl/m_news.gif" border="0">对方未读 <img src="images/uctrl/m_olds.gif" border="0">对方已读 &nbsp;&nbsp;&nbsp;&nbsp; 节省服务器空间是一种美德! 请及时删除过时信息!  &nbsp;&nbsp;
      <input type="checkbox" id="selectall" name="selallmsg" onclick="chkselall(this.form);">
      <label for="selectall">全选</label>
      &nbsp;&nbsp;
      <input type="submit" value=" 删除选择 "></td>
  </tr>
  <%else%>
  <tr>
    <td class=ctd colspan="6" height=30><b>发件箱暂时没有短信</b></td>
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
			return confirm("你准备删除选择的 " + total + " 条信息! 删除后将不能恢复,确信吗?");
		else
			{alert("请先选择您要删除的信息!");return false;}
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
