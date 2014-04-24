<!--#include file="include/conn.asp"-->
<%
'Call ChkPageAble(2)
CurPage="首页"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage="index"
Call FileInc(0, "js/login.js")
xjweb.header()
Call TopTable()
Call BottomTable()
xjweb.footer()
closeObj()
%>