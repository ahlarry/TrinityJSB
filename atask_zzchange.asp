<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'16:45 2007-2-9-星期五
Call ChkPageAble(4)
CurPage="调试任务 → 更改调试任务责任人"
strPage="atask"
'Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchLsh()%>
		</td></tr>
		<Tr><Td class=ctd height=300>
			<%Call ataskZzchange()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function ataskZzchange()
	Dim s_lsh
	s_lsh=""
	If Trim(Request("s_lsh"))<>"" Then s_lsh=Trim(Request("s_lsh"))
	If s_lsh="" Then Call TbTopic("请输入要更改责任人任务书的流水号!") : Exit Function

	strSql="select * from [mtask] where lsh='"&s_lsh&"'"
	Set Rs=xjweb.Exec(strSql,1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书不存在!","atask_zzchange.asp")
	ElseIf IsNull(Rs("sjjssj")) Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书正在设计中!","atask_zzchange.asp")
	ElseIf Rs("mjjs") Then
		Call JsAlert("流水号 【" & s_lsh & "】 任务书已经全部完成!不能更改!","atask_zzchange.asp")
	Else
		If Rs("group")=Session("userGroup") Or Rs("zz")=Session("userName") Or Rs("sjzz")=Session("userName") Or Rs("jgzz")=Session("userName") Or Session("userGroup")=5 Then
			Call atask_zzchange(rs)
		Else
			Call JsAlert("流水号为 【" & s_lsh & "】 的任务书的组长是 "&rs("zz")&"!\nSorry! 您无权更改!","atask_zzchange.asp")
		End If
	End If
	Rs.Close
End Function

function atask_zzchange(rs)
	Call mtask_fewinfo(rs)
	Response.Write(XjLine(10,"100%",""))
	Response.Write(XjLine(1,"100%",web_info(12)))
%>
	<%Call TbTopic("更改流水号 <font style=color:#0000FF>"&rs("lsh")&"</font> 的任务书")%>
	<table class=xtable cellspacing=0 cellpadding=3 width="95%">
		<form action="atask_zzchangeindb.asp" method=post>
		<tr>
		<%select case rs("mjxx")%>
			<%case "全套"%>
				<td class=ctd width="10%">模头调试单</td>
				<%if session("userGroup")=rs("group") Or Rs("sjzz")=Session("userName") Or Rs("jgzz")=Session("userName") then%>
					<td class=ctd width="100">
						<%if isnull(rs("mttsdr")) then%>
							&nbsp;
						<%else%>
							<select name=mttsdr>
								<option value=<%=rs("mttsdr")%>><%=rs("mttsdr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="100"><%=rs("mttsdr")%>&nbsp;</td>
				<%end if%>

				<%if session("userGroup")=5 then%>
					<td class=ctd width="10%">模头调试</td>
					<td class=ctd width="100">
						<%if isnull(rs("mttsr")) then%>
							&nbsp;
						<%else%>
							<select name=mttsr>
								<option value=<%=rs("mttsr")%>><%=rs("mttsr")%></option>
								<%for i=0 to ubound(c_allzy)%>
									<option value=<%=c_allzy(i)%>><%=c_allzy(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>

					<td class=ctd width="*">模头调试信息整理</td>
					<td class=ctd width="100">
						<%if isnull(rs("mttsxxzlr")) then%>
							&nbsp;
						<%else%>
							<select name=mttsxxzlr>
								<option value=<%=rs("mttsxxzlr")%>><%=rs("mttsxxzlr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="10%">模头调试</td>
					<td class=ctd width="100"><%=rs("mttsr")%>&nbsp;</td>
					<td class=ctd width="*">模头调试信息整理</td>
					<td class=ctd width="100"><%=rs("mttsxxzlr")%>&nbsp;</td>
				<%end if%>
			</tr>
			<tr>
				<%if session("userGroup")=rs("group") Or Rs("sjzz")=Session("userName") Or Rs("jgzz")=Session("userName") then%>
					<td class=ctd width="10%">定型调试单</td>
					<td class=ctd width="100">
						<%if isnull(rs("dxtsdr")) then%>
							&nbsp;
						<%else%>
							<select name=dxtsdr>
								<option value=<%=rs("dxtsdr")%>><%=rs("dxtsdr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="10%">定型调试单</td>
					<td class=ctd width="100"><%=rs("dxtsdr")%>&nbsp;</td>
				<%end if%>

				<%if session("userGroup")=5 then%>
					<td class=ctd width="10%">定型调试</td>
					<td class=ctd width="100">
						<%if isnull(rs("dxtsr")) then%>
							&nbsp;
						<%else%>
							<select name=dxtsr>
								<option value=<%=rs("dxtsr")%>><%=rs("dxtsr")%></option>
								<%for i=0 to ubound(c_allzy)%>
									<option value=<%=c_allzy(i)%>><%=c_allzy(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
					<td class=ctd width="*">定型调试信息整理</td>
					<td class=ctd width="100">
						<%if isnull(rs("dxtsxxzlr")) then%>
							&nbsp;
						<%else%>
							<select name=dxtsxxzlr>
								<option value=<%=rs("dxtsxxzlr")%>><%=rs("dxtsxxzlr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="25%">定型调试</td>
					<td class=ctd width="100"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=ctd width="*">定型调试信息整理</td>
					<td class=ctd width="100"><%=rs("dxtsxxzlr")%>&nbsp;</td>
				<%end if%>
				</tr>
			<%case "模头"%>
				<%if session("userGroup")=rs("group") Or Rs("sjzz")=Session("userName") Or Rs("jgzz")=Session("userName") then%>
					<td class=ctd width="10%">模头调试单</td>
					<td class=ctd width="100">
						<%if isnull(rs("mttsdr")) then%>
							&nbsp;
						<%else%>
							<select name=mttsdr>
								<option value=<%=rs("mttsdr")%>><%=rs("mttsdr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="10%">模头调试单</td>
					<td class=ctd width="100"><%=rs("mttsdr")%>&nbsp;</td>
				<%end if%>

				<%if session("userGroup")=5 then%>
					<td class=ctd width="10%">模头调试</td>
					<td class=ctd width="100">
						<%if isnull(rs("mttsr")) then%>
							&nbsp;
						<%else%>
							<select name=mttsr>
								<option value=<%=rs("mttsr")%>><%=rs("mttsr")%></option>
								<%for i=0 to ubound(c_allzy)%>
									<option value=<%=c_allzy(i)%>><%=c_allzy(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>

					<td class=ctd width="*">模头调试信息整理</td>
					<td class=ctd width="100">
						<%if isnull(rs("mttsxxzlr")) then%>
							&nbsp;
						<%else%>
							<select name=mttsxxzlr>
								<option value=<%=rs("mttsxxzlr")%>><%=rs("mttsxxzlr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="10%">模头调试</td>
					<td class=ctd width="100"><%=rs("mttsr")%>&nbsp;</td>
					<td class=ctd width="*">模头调试信息整理</td>
					<td class=ctd width="100"><%=rs("mttsxxzlr")%>&nbsp;</td>
				<%end if%>
			<%case "定型"%>
				<%if session("userGroup")=rs("group") Or Rs("sjzz")=Session("userName") Or Rs("jgzz")=Session("userName") then%>
					<td class=ctd width="10%">定型调试单</td>
					<td class=ctd width="100">
						<%if isnull(rs("dxtsdr")) then%>
							&nbsp;
						<%else%>
							<select name=dxtsdr>
								<option value=<%=rs("dxtsdr")%>><%=rs("dxtsdr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="10%">定型调试单</td>
					<td class=ctd width="100"><%=rs("dxtsdr")%>&nbsp;</td>
				<%end if%>

				<%if session("userGroup")=5 then%>
					<td class=ctd width="10%">定型调试</td>
					<td class=ctd width="100">
						<%if isnull(rs("dxtsr")) then%>
							&nbsp;
						<%else%>
							<select name=dxtsr>
								<option value=<%=rs("dxtsr")%>><%=rs("dxtsr")%></option>
								<%for i=0 to ubound(c_allzy)%>
									<option value=<%=c_allzy(i)%>><%=c_allzy(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
					<td class=ctd width="*">定型调试信息整理</td>
					<td class=ctd width="100">
						<%if isnull(rs("dxtsxxzlr")) then%>
							&nbsp;
						<%else%>
							<select name=dxtsxxzlr>
								<option value=<%=rs("dxtsxxzlr")%>><%=rs("dxtsxxzlr")%></option>
								<%for i=0 to ubound(c_jsb)%>
									<option value=<%=c_jsb(i)%>><%=c_jsb(i)%></option>
								<%next%>
							</select>
						<%end if%>
					</td>
				<%else%>
					<td class=ctd width="10%">定型调试</td>
					<td class=ctd width="100"><%=rs("dxtsr")%>&nbsp;</td>
					<td class=ctd width="*">定型调试信息整理</td>
					<td class=ctd width="100"><%=rs("dxtsxxzlr")%>&nbsp;</td>
				<%end if%>
				</tr>
			<%case else%>
				<%response.write(rs("mjxx"))%>
		<%end select%>
		</tr>
		<tr><td class=ctd colspan=8><input type=submit value=" 更改 "></td></tr>
		<input type=hidden name=lsh value=<%=rs("lsh")%>>
		</form>
	</table>

<%
end function
%>