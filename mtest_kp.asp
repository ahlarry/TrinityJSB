<!--#include file="include/conn.asp"-->
<!--#include file="include/page/user_dbinf.asp"-->
<%
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="模具调试 → 调试考评列表"
strPage="mtest"
'Call FileInc(0, "js/ftask.js")
xjweb.header()
Call TopTable()

'定义变量及变量赋值
Dim iyear, imonth, dtstart, dtend, struser, irwzf, iaddfz, ilxrwzf, icount
iyear = request("searchy")
imonth = request("searchm")
struser = request("searchuser")
If iyear = "" Then iyear = year(now)
If imonth = "" Then imonth = month(now)

dtend=cdate(iyear&"年"&imonth&"月1日")
dtend=dateadd("m",1,dtend)
dtend=dateadd("d",-1,dtend)
dtstart=cdate(iyear&"年"&imonth&"月1日")

Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
	<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
		<Tr><Td class=ctd>
			<%Call SearchMantime()%>
		</Td></Tr>
		<Tr><Td class=ctd height=300>
			<%Call mtestList()%>
			<%Response.Write(XjLine(10,"100%",""))%>
		</Td></Tr>
	</Table>
<%
End Sub

Function SearchMantime()
%>
	<table cellpadding=2 cellspacing=0>
		<form action=<%=request.servervariables("script_name")%> method=get>
		<tr>
			<td>
			请选择:
			<select name="searchy" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value);'>
				<%for i = year(now) - 3 to year(now)%>
					<option value=<%=i%><%If i = cint(iyear) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>年
			<select name="searchm" onchange='location.href("<%=request.servervariables("script_name")%>?searchy="+this.form.searchy.value+"&searchm="+this.form.searchm.value);'>
				<%for i = 1 to 12%>
					<option value=<%=i%><%If i = cint(imonth) Then%> selected<%end If%>><%=i%></option>
				<%next%>
			</select>月&nbsp;&nbsp;
			<input type="submit" value=" 选 择 ">
			</td>
		</tr>
		</form>
	</table>
<%
End Function

Function mtestList()
	Call TbTopic("调试完成信息")
	Dim Tmplsh,hgs,hgf,hgz,cts,ctf,ctz,jts,jtf,jtz,jys,jyf,jyz,lcs,lcf,lcz
	Tmplsh="" : hgs=0 : hgf=0 : hgz=0 : cts=0 : ctf=0 : ctz=0 : jts=0 : jtf=0 : jtz=0 : jys=0 : jyf=0 : jyz=0 : lcs=0 : lcf=0 : lcz=0
	strSql="select * from [mantime] where zrr='TT调试员' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Do While Not Rs.eof
		If Tmplsh<>Rs("lsh") Then
			select case Mid(Rs("rwlr"),3)
				case "厂内初调"
					cts=cts+1
					If Rs("jc")>1 Then ctz=ctz+1
					If Rs("jc")<1 Then ctf=ctf+1
				case "厂外精调"
					jts=jts+1
					If Rs("jc")>1 Then jtz=jtz+1
					If Rs("jc")<1 Then jtf=jtf+1
				case "预验收或寄样"
					jys=jys+1
					If Rs("jc")>1 Then jyz=jyz+1
					If Rs("jc")<1 Then jyf=jyf+1
				case "来厂验收"
					lcs=lcs+1
					If Rs("jc")>1 Then lcz=lcz+1
					If Rs("jc")<1 Then lcf=lcf+1
				case else
					hgs=hgs+1
					If Rs("jc")>1 Then hgz=hgz+1
					If Rs("jc")<1 Then hgf=hgf+1
			end select
			Tmplsh=Rs("lsh")
		End If
		Rs.movenext
	Loop
	Rs.Close
%>
<table width="80%" cellpadding=2 cellspacing=0 class="xtable" align="center">
  <tr>
    <th class=th>&nbsp;</th>
    <th class=th>调试合格</th>
    <th class=th>厂内初调</th>
    <th class=th>厂外精调</th>
    <th class=th>验收寄样</th>
    <th class=th>来厂验收</th>
    <th class=th>合计有效套数</th>
  </tr>
  <tr>
    <td class=ctd>多余额定考核次数</th>
    <td class=ctd><%=hgf%></td>
    <td class=ctd><%=ctf%></td>
    <td class=ctd><%=jtf%></td>
    <td class=ctd><%=jyf%></td>
    <td class=ctd><%=lcf%></td>
    <td class=ctd><%=hgf+ctf+jtf+jyf+lcf%></td>
  </tr>
  <tr>
    <td class=ctd>少于额定考核次数</th>
    <td class=ctd><%=hgz%></td>
    <td class=ctd><%=ctz%></td>
    <td class=ctd><%=jtz%></td>
    <td class=ctd><%=jyz%></td>
    <td class=ctd><%=lcz%></td>
    <td class=ctd><%=hgz+ctz+jtz+jyz+lcz%></td>
  </tr>
  <tr>
    <td class=ctd>合计(包括考核范围内模具)</td>
    <td class=ctd><%=hgs%></td>
    <td class=ctd><%=cts%></td>
    <td class=ctd><%=jts%></td>
    <td class=ctd><%=jys%></td>
    <td class=ctd><%=lcs%></td>
    <td class=ctd><%=hgs+cts+jts+jys+lcs%></td>
  </tr>
</table>
<%Call TbTopic("分值统计")
	Dim sjjl,sjcf
	sjjl=0 : sjcf=0
  	strSql="select * from [mantime] where rwlr like '%调试合格(%' and datediff('d',jssj,'"&dtstart&"')<=0 and datediff('d',jssj,'"&dtend&"')>=0 order by lsh desc"
	Set Rs=xjweb.Exec(strSql, 1)
	Do While Not Rs.eof
		If Rs("fz")>0 Then
			sjjl=sjjl+Rs("fz")
		else
			sjcf=sjcf+Rs("fz")
		End If
		Rs.movenext
	Loop
	Rs.Close

	Dim sjzzjl,sjzzcf,tszzjl,tszzcf,xxzzjl,xxzzcf,gfwhjl,gfwhcf,jsdbjl,jsdbcf
	sjzzjl=0 : sjzzcf=0 : tszzjl=0 : tszzcf=0 : xxzzjl=0
	xxzzcf=0 : gfwhjl=0 : gfwhcf=0 : jsdbjl=0 : jsdbcf=0
  	strSql="select * from [kp_jsb] where kp_item like '%于额定次数%' and datediff('d',kp_time,'"&dtstart&"')<=0 and datediff('d',kp_time,'"&dtend&"')>=0"
	Set Rs=xjweb.Exec(strSql, 1)
	Do While Not Rs.eof
		select case Rs("kp_zrrjs")
			case "设计组长"
				If Rs("kp_item")="调试少于额定次数" Then
					sjzzjl=sjzzjl+Rs("kp_uprice")
				else
					sjzzcf=sjzzcf-Rs("kp_uprice")
				End If
			case "调试组长"
				If Rs("kp_item")="调试少于额定次数" Then
					tszzjl=tszzjl+Rs("kp_uprice")
				else
					tszzcf=tszzcf-Rs("kp_uprice")
				End If
			case "信息组长"
				If Rs("kp_item")="调试少于额定次数" Then
					xxzzjl=xxzzjl+Rs("kp_uprice")
				else
					xxzzcf=xxzzcf-Rs("kp_uprice")
				End If
			case "规范维护"
				If Rs("kp_item")="调试少于额定次数" Then
					gfwhjl=gfwhjl+Rs("kp_uprice")
				else
					gfwhcf=gfwhcf-Rs("kp_uprice")
				End If
			case "技术代表"
				If Rs("kp_item")="调试少于额定次数" Then
					jsdbjl=jsdbjl+Rs("kp_uprice")
				else
					jsdbcf=jsdbcf-Rs("kp_uprice")
				End If
		end select
		Rs.movenext
	Loop
	Rs.Close

	sjzzjl=Round(sjzzjl,1)
	sjzzcf=Round(sjzzcf,1)
	tszzjl=Round(tszzjl,1)
	tszzcf=Round(tszzcf,1)
	xxzzjl=Round(xxzzjl,1)
	xxzzcf=Round(xxzzcf,1)
	gfwhjl=Round(gfwhjl,1)
	gfwhcf=Round(gfwhcf,1)
	jsdbjl=Round(jsdbjl,1)
	jsdbcf=Round(jsdbcf,1)
%>
<table width="80%" cellpadding=2 cellspacing=0 class="xtable" align="center">
  <tr>
    <th class=th>&nbsp;</th>
    <th class=th>设计人员</th>
    <th class=th>设计组长</th>
    <th class=th>调试组长</th>
    <th class=th>信息组长</th>
    <th class=th>规范维护</th>
    <th class=th>技术代表</th>
  </tr>
  <tr>
    <td class=ctd>奖励</td>
    <td class=ctd><%=sjjl%></td>
    <td class=ctd><%=sjzzjl%></td>
    <td class=ctd><%=tszzjl%></td>
    <td class=ctd><%=xxzzjl%></td>
    <td class=ctd><%=gfwhjl%></td>
    <td class=ctd><%=jsdbjl%></td>
  </tr>
  <tr>
    <td class=ctd>处罚</th>
    <td class=ctd><%=sjcf%></td>
    <td class=ctd><%=sjzzcf%></td>
    <td class=ctd><%=tszzcf%></td>
    <td class=ctd><%=xxzzcf%></td>
    <td class=ctd><%=gfwhcf%></td>
    <td class=ctd><%=jsdbcf%></td>
  </tr>
  <tr>
    <td class=ctd>合计</th>
    <td class=ctd><%=sjjl+sjcf%></td>
    <td class=ctd><%=sjzzjl+sjzzcf%></td>
    <td class=ctd><%=tszzjl+tszzcf%></td>
    <td class=ctd><%=xxzzjl+xxzzcf%></td>
    <td class=ctd><%=gfwhjl+gfwhcf%></td>
    <td class=ctd><%=jsdbjl+jsdbcf%></td>
  </tr>
</table>
<%
end function
%>