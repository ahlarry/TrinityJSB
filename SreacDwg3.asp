<!--#include file="include/conn.asp"-->
<!--#include file="easyasp/easp.asp" -->
<%
CurPage="断面图和投影图"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage=""
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, strlsh
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder
Call Main()
Call BottomTable()
xjweb.footer()
closeObj()

Sub Main()
%>
<Table class=xtable cellspacing=0 cellpadding=0 width="<%=web_info(8)%>">
<form name="form1" method="post" action="<% =Request.ServerVariables("PATH_INFO")%>">
请输入模具流水号:
<input name="keyword" type="text" id="keyword">
<input type="submit" name="Submit" value="搜索">
</form>
</Table>
<%
Dim newsearch, s_lsh, action, keyword
keyword=""
s_lsh=Trim(Request("s_lsh"))
If s_lsh<>"" Then
	keyword=s_lsh
Else
	keyword=Request.Form("keyword")
End If
if keyword<>"" then
Set newsearch=new SearchFile
'newsearch.Folders=Server.mappath("dmtj")
'newsearch.Folders="G:\设计参考\断面图集" '是绝对路径
newsearch.Folders="D:\模具图" '是绝对路径
newsearch.keyword=keyword
newsearch.Search
Set newsearch=Nothing
'Set newsearch=new SearchFile
'newsearch.Folders="F:\模具图档\模具修理" '是绝对路径
'newsearch.keyword=keyword2
'newsearch.Search
'Set newsearch=Nothing
'response.Write("<br/>费时："&(timer()-st)*1000&"毫秒")
end if
End Sub
%>