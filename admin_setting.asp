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
	Call TbTopic(web_info(0) & " ϵͳ����")
%>
	<table cellspacing=0 cellpadding=3 class=xtable width="60%">
	<form action="<%=request.servervariables("script_name")%>?action=setting" method="post" onsubmit="return chkinf();">
		<tr>
			<td class=rtd width="20%">��վ����</td>
			<td class=ltd width="*"><input type="text" size="50" name="setting0" value="<%=web_info(0)%>"></td>
			<td class=ctd width="10%"><img src="images/admin/help.gif" alt="��վ������"></td>
		</tr>
		<tr>
			<td class=rtd>�汾��</td>
			<td class=ltd><input type="text" size="50" name="setting1" value="<%=web_info(1)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="����İ汾��"></td>
		</tr>
		<tr>
			<td class=rtd>����Ŀ¼</td>
			<td class=ltd><input type="text" size="50" name="setting2" value="<%=web_info(2)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="ϵͳ������Ŀ¼()"></td>
		</tr>
		<tr>
			<td class=rtd>��վ����</td>
			<td class=ltd><input type="text" size="50" name="setting3" value="<%=web_info(3)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��վ������"></td>
		</tr>
		<tr>
			<td class=rtd>��վ����</td>
			<td class=ltd><input type="text" size="50" name="setting4" value="<%=web_info(4)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��վ����"></td>
		</tr>
		<tr>
			<td class=rtd>��վ�ؼ���</td>
			<td class=ltd><input type="text" size="50" name="setting5" value="<%=web_info(5)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��վ�ؼ���"></td>
		</tr>
		<tr>
			<td class=rtd>�Ҽ�����</td>
			<td class=ltd><input type="checkbox" id="brm" name="setting6" value="true" <%if web_info(6) then%>checked<%end if%>><label for="brm">����</label></td>
			<td class=ctd><img src="images/admin/help.gif" alt="�Ƿ������Ҽ�����"></td>
		</tr>
		<tr>
			<td class=rtd>״̬��</td>
			<td class=ltd><input type="text" size="50" name="setting7" value="<%=web_info(7)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��ҳ��״̬����ʾ����(��Ϊ����ʾ)"></td>
		</tr>
		<tr>
			<td class=rtd>��ҳ���</td>
			<td class=ltd><input type="text" size="50" name="setting8" value="<%=web_info(8)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��ҳ�Ŀ��"></td>
		</tr>
		<tr>
			<td class=rtd>��վ��LOGO</td>
			<td class=ltd><input type="text" size="50" name="setting9" value="<%=web_info(9)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��վ��LOGO"></td>
		</tr>
		<tr>
			<td class=rtd>Cookie����</td>
			<td class=ltd><input type="text" size="50" name="setting10" value="<%=web_info(10)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��վCookie,���ظ�ʱ�����"></td>
		</tr>
		<tr>
			<td class=rtd>Ȩ��˵��</td>
			<td class=ltd><input type="text" size="50" name="setting11" value="<%=web_info(11)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="Ȩ��˵��"></td>
		</tr>
		<tr>
			<td class=rtd>��ҳ����ɫ</td>
			<td class=ltd><input type="text" size="40" name="setting12" value="<%=web_info(12)%>" onblur="document.all.tttccc.style.color=this.value;">&nbsp;&nbsp;<font style=color:<%=web_info(12)%> id="tttccc">������</font></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��ҳ�Ŀ��"></td>
		</tr>		
		<tr>
			<td class=rtd>��վ��Ȩ</td>
			<td class=ltd><input type="text" size="50" name="setting13" value="<%=web_info(13)%>"></td>
			<td class=ctd><img src="images/admin/help.gif" alt="��վ��Ȩ"></td>
		</tr>
		<tr>
			<td class=rtd>&nbsp;</td>
			<td class=ctd><input type=submit value="��������"></td>
			<td class=ctd>&nbsp;</td>
		</tr>
		</form>
	</table>
	<script language="javascript">
		function chkinf()
			{
				var da=document.all;
				if(da.setting0.value==""){alert('��վ����Ϊ��!');da.setting0.focus();return false;}
				if(da.setting1.value==""){alert('�汾��Ϊ��!');da.setting1.focus();return false;}
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
		vbcrlf & "web_info(0)="""&strSiteName&"""						'��վ����" &_
		vbcrlf & "web_info(1)="""&strVersion&"""							'�汾��" &_
		vbcrlf & "web_info(2)="""&strVirtualDir&"""					'����Ŀ¼" &_
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
  Call JsAlert("ϵͳ��Ϣ���ĳɹ�!","?action=main")
End Function
%>