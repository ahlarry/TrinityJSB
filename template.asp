<!--#include file="include/conn.asp"-->
<%
'Call ChkPageAble(2)
CurPage="��ҳ"					'ҳ�������λ��( ��������� �� ���������)
strPage="index"
Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Call BottomTable()
xjweb.footer()
closeObj()
%>