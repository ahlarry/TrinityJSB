<!--#include file="include/conn.asp"-->
<%
'Call ChkPageAble(2)
CurPage="��ҳ"
strPage=""  'index
Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Rem ����Ϊ���ļ��ĺ�������
Sub Main()
%>
	<Table cellpadding=0 cellspacing=0 width="<%=web_info(8)%>">
		<tr>
		<td width=180 valign=top>
			<%Call UserLogin()%>
			<%Response.Write(xjLine(2,"100%",""))%>
			<%=TaskStat()%>
			<%Response.Write(xjLine(2,"100%",""))%>
			<%=AuthorInfo()%>
		</td>
		<td width=2></td>
		<td width=* valign=top>
			<%=IndexCenter()%>
			<%=DayReciteWords(5)%>
		</td>
		<td width=2></td>
		<td width=180 valign=top>
			<%=CurTask()%>
			<%Response.Write(xjLine(2,"100%",""))%>
			<%=CurTestMould()%>
			<%Response.Write(xjLine(2,"100%",""))%>
			<%=OnlineUser()%>
		</td>
		</tr>
	</table>
<%
End Sub
Function UserLogin()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> ���û���¼</td></tr>
		<Tr height=100><td class=ctd><%=FastLogin()%></td></tr>
	</Table>
<%
End Function
Function TaskStat()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> ��������ͳ��</td></tr>
		<Tr height=80><td class=ltd>
				<li>����������<%=xjweb.rscount("mtask")%> ��<br></li>
				<li>ȥ������<%=xjweb.rscount("mtask where Year(rwxdsj)=(Year(now)-1)")%> ��<br></li>
				<li>����Ŀǰ��<%=xjweb.rscount("mtask where Year(rwxdsj)=Year(now)")%> ��<br></li>
				<li>������ƣ�<%=xjweb.rscount("mtask where isnull(sjjssj)")%> ��</li>
		</td></tr>
	</Table>
<%
End Function

Function AuthorInfo()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> ��������Ϣ</td></tr>
		<Tr><td class=ltd>
			<b>My MSN:</b><br>
				<span style=padding:10; alt="�����ҵ�MSN�����³���ϵ��">ahxujian@hotmail.com<br></span>
			<b>My QQ:</b><br>
				<span style=padding:10; alt="�����ҵ�QQ��û�³���ϵ��">119891935<br></span>
			<b>My Email:</b><br>
				<a href=mailto:xujian@chinatrinity.com ><span  style=padding:10; alt="�����ҵ�Email������û�³���ϵ��">xujian@chinatrinity.com</span></a>
		</td></tr>
	</Table>
<%
End Function

Function UserStat()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> ���û�ͳ��</td></tr>
		<Tr><td class=ltd>
			�û�������<%=xjweb.rscount("[ims_user]")%> ��<br>
			�鳤��<%=xjweb.rscount("[ims_user] where Mid(user_able,4,1)=1")%> ��<br>
			��Ա��<%=xjweb.rscount("[ims_user] where Mid(user_able,5,1)=1")%> ��<br>
		</td></tr>
	</Table>
<%
End Function
Function CurTask()
%>
	<Table class=xtable cellpadding=2 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> �������������</td></tr>
		<Tr><td class=ltd>
	<%
		Set Rs=xjweb.Exec("Select top 5 lsh, dmmc, dwmc from [mtask] where isnull(sjjssj) order by lsh desc",1)
		i=1
		Do While Not Rs.Eof
			Response.Write("<li><a href=mtask_list.asp?lsh="&rs("lsh")&" alt=""��ˮ��:"&rs("lsh")&"<br>��λ����:"&rs("dwmc")&"<br>��������:"&rs("dmmc")&""">" & xjweb.StringCut(rs("dwmc") & "--" & rs("dmmc"), 15) & "</a></li>")
			Rs.moveNext
			i=i+1
		Loop
		Rs.Close
	%>
		<span style=text-align:right;width:'100%';padding-right:10;><a href=mtask_list.asp?term=no>more...</a></span>
		</td></tr>
	</Table>
<%
End Function
Function CurTestMould()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> �����µ���ģ��</td></tr>
		<Tr><td class=ltd>
	<%
		Set Rs=xjweb.Exec("Select top 5 a.lsh as lsh, dmmc, dwmc from [mtask] a, [ts_mould] b where a.lsh=b.lsh and isnull(tsjssj) order by tsgxsj desc",1)
		i=1
		Do While Not Rs.Eof
			Response.Write("<li><a href=mtest_list.asp?lsh="&rs("lsh")&" alt=""��ˮ��:"&rs("lsh")&"<br>��λ����:"&rs("dwmc")&"<br>��������:"&rs("dmmc")&""">" & xjweb.StringCut(rs("dwmc") & "--" & rs("dmmc"), 12) & "</a></li>")
			Rs.moveNext
			i=i+1
		Loop
		Rs.Close
	%>
		<span style=text-align:right;width:'100%';padding-right:10;><a href=mtest_list.asp?term=no>more...</a></span>
		</td></tr>
	</Table>
<%
End Function
Function OnlineUser()
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th style=text-align:left;> 
		������ͳ��</td></tr>
		<Tr><td class=ltd>
			<li alt="��2005��2��15������">ϵͳ��������: <%=xjweb.rscount("[ims_stat]")%> ��</li>
			<li>��ǰ�����û�: <a href="online.asp?action=list"><%=xjweb.rscount("[ims_online]")%></a> ��</li>
			<li>������ʴ���: <%=xjweb.rscount("[ims_stat] where datediff('d',stat_time,'"&now()&"')=1")%> ��</li>
			<li>������ʴ���: <%=xjweb.rscount("[ims_stat] where datediff('d',stat_time,'"&now()&"')=0")%> ��</li>
		</td></tr>
	</Table>
<%
End Function

Function indexCenter()
	strSql="select user_name, user_birthday from [ims_user] where Month([user_birthday])=Month(now()) and Day([user_birthday])=Day(now()) order by user_name"
	Set Rs=xjweb.Exec(strSql, 1)
	If Rs.Bof Or Rs.Eof Then
		'Call indexMain()
	Else
		Call HappyBirthday(Rs)
	End If
	Rs.Close
End Function

Function indexMain()
%>
	<Table border=0 cellpadding=5 cellspacing=0 width="100%">
		<Tr><Td align=center valign=middle height=180> 
			<div style='width:300;filter:glow(color=green,strength=5);font-size:18pt;color:white;'>���Ѽ�ģ��Ϣ����ϵͳ</div><br>
			<div style='width:300;filter:glow(color=blue,strength=5);font-size:11pt;color:white;'>Ver: <%=web_info(1)%></div><br>
			<div style='width:300;filter:glow(color=red,strength=3);font-size:14pt;color:white;'>��ӭʹ��</div>
		</Td></Tr>
	</Table>
<%
End Function

Function HappyBirthday(Rs)
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th> ����ף�� </td></tr>
		<Tr><td class=ctd><img src="<%=web_info(2)%>images/birthday.jpg" onload="if(this.width>280) this.width=280;if(this.height>280) this.height=280;"><br>
		<%
			Do while not Rs.Eof
				Response.Write("<div alt=""��������:" &xjdate(rs("user_birthday"),1)&"""><font style=font-size:16pt;font-weight:bold;>" & Rs("user_name") & "</font> <font style=font-size:14pt;>���տ���!</font></div>")
				Rs.moveNext
			Loop
		%>
		</td></tr>
	</Table>
<%
End Function

Function DayReciteWords(iCount)
%>
	<Table class=xtable cellpadding=5 cellspacing=0 width="100%">
		<Tr><td class=th colspan=4> ���챳���� </td></tr>
		<%
		Dim Myvalue
		i=1
		Randomize '��ʼ���������������
		Myvalue = Int((6703 * Rnd) + 1)' �����������
		strSql="select * from [ims_words] where id > "& Myvalue &" and id <= "& Myvalue+iCount &" "
		Set Rs=xjweb.Exec(strSql,1)
		Do While Not Rs.Eof
		%>
			<Tr>
				<Td class=ctd>No.<b><%=i%></b></Td>
				<Td class=ctd><font color = #ff0000><b><%=Rs("word")%></b></font></Td>
				<Td class=ctd><font color=green><b>[<font face="Kingsoft Phonetic"><%=Rs("phonetic")%></font>]</b></font></Td>
				<Td class=ltd><font color=blue><%=Rs("Remark")%></font></Td>
			</Tr>
		<%
			i=i+1
			Rs.movenext
		Loop
		Rs.Close
		%>
	</Table>
<%
End Function
%>

