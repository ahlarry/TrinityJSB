<!--#include file="include/conn.asp"-->
<%
Call ChkAdminAble()
Dim action
action=LCase(Request("action"))
select case action
	case "setting"
		call indb()
	case else
		call main()
end select

Function main()
	Call xjweb.Header()
	Call TbTopic(web_info(0) & " 系统设置")
%>
	<table cellspacing=0 cellpadding=3 class=xtable width="60%">
	<form action="<%=request.servervariables("script_name")%>?action=setting" method="post" onsubmit="return chkinf();">
		<tr>
			<td class=rtd width="20%">网站名称</td>
			<td class=ltd width="*"><input type="text" size="50" name="setting0" value="<%=web_info(0)%>"></td>
			<td class=ctd width="10%"><img src="images/admin/help.gif" alt="网站的名称"></td>
		</tr>
		<tr>
			<td class=rtd>版本号</td>
			<td class=ltd><input type="text" size="50" name="setting1" value="<%=web_info(1)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="软件的版本号"></td>
		</tr>
		<tr>
			<td class=rtd>虚拟目录</td>
			<td class=ltd><input type="text" size="50" name="setting2" value="<%=web_info(2)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="系统的虚拟目录()"></td>
		</tr>
		<tr>
			<td class=rtd>网站作者</td>
			<td class=ltd><input type="text" size="50" name="setting3" value="<%=web_info(3)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网站的作者"></td>
		</tr>
		<tr>
			<td class=rtd>网站描述</td>
			<td class=ltd><input type="text" size="50" name="setting4" value="<%=web_info(4)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网站描述"></td>
		</tr>
		<tr>
			<td class=rtd>网站关键字</td>
			<td class=ltd><input type="text" size="50" name="setting5" value="<%=web_info(5)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网站关键字"></td>
		</tr>
		<tr>
			<td class=rtd>右键操作</td>
			<td class=ltd><input type="checkbox" id="brm" name="setting6" value="true" <%if web_info(6) then%>checked<%end if%>><label for="brm">允许</label></td>
			<td class=ctd><img src="images/admin/help.gif" alt="是否允许右键操作"></td>
		</tr>
		<tr>
			<td class=rtd>状态栏</td>
			<td class=ltd><input type="text" size="50" name="setting7" value="<%=web_info(7)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网页的状态栏显示内容(空为不显示)"></td>
		</tr>
		<tr>
			<td class=rtd>网页宽度</td>
			<td class=ltd><input type="text" size="50" name="setting8" value="<%=web_info(8)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网页的宽度"></td>
		</tr>
		<tr>
			<td class=rtd>网站主LOGO</td>
			<td class=ltd><input type="text" size="50" name="setting9" value="<%=web_info(9)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网站主LOGO"></td>
		</tr>
		<tr>
			<td class=rtd>Cookie名称</td>
			<td class=ltd><input type="text" size="50" name="setting10" value="<%=web_info(10)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网站Cookie,有重复时请更名"></td>
		</tr>
		<tr>
			<td class=rtd>权限说明</td>
			<td class=ltd><input type="text" size="50" name="setting11" value="<%=web_info(11)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="权限说明"></td>
		</tr>
		<tr>
			<td class=rtd>网页主导色</td>
			<td class=ltd><input type="text" size="40" name="setting12" value="<%=web_info(12)%>" onblur="document.all.tttccc.style.color=this.value;">&nbsp;&nbsp;<font style=color:<%=web_info(12)%> id="tttccc">███</font></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网页的宽度"></td>
		</tr>		
		<tr>
			<td class=rtd>网站版权</td>
			<td class=ltd><input type="text" size="50" name="setting13" value="<%=web_info(13)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="网站版权"></td>
		</tr>
		<tr>
			<td class=rtd>&nbsp;</td>
			<td class=ctd><input type=submit value="更改设置"></td>
			<td class=ctd>&nbsp;</td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		function chkinf()
			{
				var da=document.all;
				if(da.setting0.value==""){alert('网站名称为空!');da.setting0.focus();return false;}
				if(da.setting1.value==""){alert('版本号为空!');da.setting1.focus();return false;}
			}
	</script>
<%
	Call xjweb.Footer()
End Function

Function indb()
	Dim file_name, strFile
	file_name="include/const.asp"
	Dim strSiteName, strVersion, strVirtualDir, strAuthor, strDescription, strKeywords, bMouseRight, strStatus, strPageWidth
	Dim strLogo, strCookie, strAbleContent, strMainColor, strCopyRight
	strSiteName="" : strVersion="" : strVirtualDir="" : strAuthor="" : strDescription="" : strKeywords=""
	bMouseRight=False : strStatus="" : strPageWidth="" : strLogo="" : strCookie="" : strAbleContent=""
	strMainColor="" : strCopyRight=""
	strSiteName=Trim(Request("setting0"))
	strVersion=Trim(Request("setting1"))
	strVirtualDir=Trim(Request("setting2"))
	strAuthor=Trim(Request("setting3"))
	strDescription=Trim(Request("setting4"))
	strKeywords=Trim(Request("setting5"))
	If Request("setting6") Then bMouseRight=True
	strStatus=Trim(Request("setting7"))
	strPageWidth=Trim(Request("setting8"))
	strLogo=Trim(Request("setting9"))
	strCookie=Trim(Request("setting10"))
	strAbleContent=Trim(Request("setting11"))
	strMainColor=Trim(Request("setting12"))
	strCopyRight=Trim(Request("setting13"))

	strFile="<" & chr(37) &_
		vbcrlf & "Dim web_info(30)" &_
		vbcrlf & "web_info(0)="""&strSiteName&"""						'网站名称" &_
		vbcrlf & "web_info(1)="""&strVersion&"""							'版本号" &_
		vbcrlf & "web_info(2)="""&strVirtualDir&"""					'虚拟目录" &_
		vbcrlf & "web_info(3)="""&strAuthor&"""" &_
		vbcrlf & "web_info(4)="""&strDescription&"""" &_
		vbcrlf & "web_info(5)="""&strKeywords&"""" &_
		vbcrlf & "web_info(6)="&bMouseRight&"" &_
		vbcrlf & "web_info(7)="""&strStatus&"""" &_
		vbcrlf & "web_info(8)="""&strPageWidth&"""" &_
		vbcrlf & "web_info(9)="""&strLogo&"""" &_
		vbcrlf & "web_info(10)="""&strCookie&"""" &_
		vbcrlf & "web_info(11)="""&strAbleContent& """" &_
		vbcrlf & "web_info(12)="""&strMainColor&"""" &_
		vbcrlf & "web_info(13)="""&strCopyRight&"""" &_
		vbcrlf & chr(37) & ">"
  Call create_file(file_name,strFile)
  Call JsAlert("系统信息更改成功!","?action=main")
End Function
%>