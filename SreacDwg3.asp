<!--#include file="include/conn.asp"-->
<!--#include file="easyasp/easp.asp" -->
<%
CurPage="����ͼ��ͶӰͼ"					'ҳ�������λ��( ��������� �� ���������)
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
������ģ����ˮ��:
<input name="keyword" type="text" id="keyword">
<input type="submit" name="Submit" value="����">
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
'newsearch.Folders="G:\��Ʋο�\����ͼ��" '�Ǿ���·��
newsearch.Folders="D:\ģ��ͼ" '�Ǿ���·��
newsearch.keyword=keyword
newsearch.Search
Set newsearch=Nothing
'Set newsearch=new SearchFile
'newsearch.Folders="F:\ģ��ͼ��\ģ������" '�Ǿ���·��
'newsearch.keyword=keyword2
'newsearch.Search
'Set newsearch=Nothing
'response.Write("<br/>��ʱ��"&(timer()-st)*1000&"����")
end if
End Sub
%>