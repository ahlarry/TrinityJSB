<!--#include file="include/conn.asp"-->
<!--#include file="include/page/mtaskinfo.asp"-->
<%
'10:37 2011-12-07
Call ChkPageAble(0)
Call ChkDepart("技术部")
CurPage="设计任务 → 我的任务"
strPage="mtask"
xjweb.header()
Call TopTable()
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()

%>
<table class="xtable" cellspacing="0" cellpadding="2" width="<%=web_info(8)%>">
  <tr>
    <td class="ctd" height="280"><%Call TaskList()%>
      <%Response.Write(XjLine(10,"100%",""))%></td>
  </tr>
</table>
<%
End Sub

Function TaskList()
	Call TbTopic(session("userName") & "的设计任务")
%>
	<table width="98%" cellpadding="2" cellspacing="0" border="0"  class="xtable"  align="center">
  <tr>
    <th class="th" width="25">id
      </td>
    </th>
    <th class="th" width="80">流水号
      </td>
    </th>
    <th class="th" width="80">任务内容
      </td>
    </th>
    <th class="th" width="80">结构
      </td>
    </th>
    <th class="th" width="80">设计
      </td>
    </th>
    <th class="th" width="80">审核
      </td>
    </th>
    <th class="th" width="*">计划结构完成
      </td>
    </th>
    <th class="th" width="100">计划全套完成
      </td>
    </th>
  </tr>
  <%
  	Dim strjg, strsj, strsh, struser, jgjgsj, sjjgsj, Tmpsj, ijgsj, isj
  	struser=session("userName") : i=1
	strSql = "select * from [mtask] where isnull(sjjssj) and (mtjgr='"&struser&"' or dxjgr='"&struser&"' or gjjgr='"&struser&"' or mtsjr='"&struser&"' or dxsjr='"&struser&"' or gjsjr='"&struser&"' or mtjgshr='"&struser&"' or dxjgshr='"&struser&"' or gjjgshr='"&struser&"' or mtsjshr='"&struser&"' or dxsjshr='"&struser&"' or gjsjshr='"&struser&"' or mtshr='"&struser&"' or dxshr='"&struser&"' or gjshr='"&struser&"')"
	set rs=xjweb.Exec(strSql, 1)
	Do while not rs.eof
		strjg=0 : strsj=0 : strsh=0
		If (Rs("mtjgr")=struser and isNull(Rs("mtjgjs"))) or (Rs("dxjgr")=struser and isNull(Rs("dxjgjs"))) or (Rs("gjjgr")=struser and isNull(Rs("gjjgjs"))) Then
	  		strjg=1
		End If
		If (Rs("mtsjr")=struser and isNull(Rs("mtsjjs"))) or (Rs("dxsjr")=struser and isNull(Rs("dxsjjs"))) or (Rs("gjsjr")=struser and isNull(Rs("gjsjjs"))) Then
	  		strsj=1
		End If
		If  (Rs("mtjgshr")=struser and isNull(Rs("mtjgshjs"))) or (Rs("dxjgshr")=struser and isNull(Rs("dxjgshjs"))) or (Rs("gjjgshr")=struser and isNull(Rs("gjjgshjs"))) or (Rs("mtsjshr")=struser and isnull(Rs("mtsjshjs"))) or (Rs("dxsjshr")=struser and isnull(Rs("dxsjshjs"))) or (Rs("gjsjshr")=struser and isnull(Rs("gjsjshjs"))) or (Rs("mtshr")=struser and isnull(Rs("mtshjs"))) or (Rs("dxshr")=struser and isnull(Rs("dxshjs"))) or (Rs("gjshr")=struser and isnull(Rs("gjshjs"))) Then
	  		strsh=1
		End If
		If strjg+strsj+strsh>0 Then
If rs("jhkssj")<>"" Then
	Tmpsj=rs("jhkssj")
else
	Tmpsj=rs("rwxdsj")
End If
  jgjgsj=datediff("d", Tmpsj, rs("jhjssj"))/2
  sjjgsj=datediff("d", Tmpsj, rs("jhjssj"))+1-jgjgsj
  isj=datediff("h", Tmpsj, rs("jhjssj"))+24
	If isj>24 Then
		ijgsj=(isj-24)/2
	else
		isj=24
		ijgsj=24
	End If
	%>
 	<tr>
    	<td class="ctd"><%=i%></td>
    	<td class="ctd"><a href="mtask_display.asp?s_lsh=<%=rs("lsh")%>"><%=rs("lsh")%></a></td>
    	<td class="ctd"><%=rs("mjxx") &  rs("rwlr")%></td>
    	<td class="ctd"><%if strjg=1 Then Response.Write("√")%>&nbsp;</td>
    	<td class="ctd"><%if strsj=1 Then Response.Write("√")%>&nbsp;</td>
    	<td class="ctd"><%if strsh=1 Then Response.Write("√")%>&nbsp;</td>
    	<td class="ctd"><%=Dateadd("h",ijgsj+1,rs("jhkssj"))%></td>
    	<td class="ctd"><%=rs("jhjssj")%></td>
	<tr>
 	<%
		i = i + 1
		End if
		rs.movenext
	loop
	%>
</table>
<%
End Function
%>
