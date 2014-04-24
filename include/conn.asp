<%@ LANGUAGE = "VBScript" CodePage = "936"%>
<%
Option Explicit
Response.Expires=0
Response.Buffer=True
Session.Timeout=120
Response.Expires = -1
Response.CacheControl="no-chache"
Response.ExpiresAbsolute= Now() - 100

Dim StartTime, CssFiles, JsFiles, ErrInfo, IsDebug, strCode, CurPage, strPage, i, strSql, strMsg
StartTime=Timer() : CssFiles="" : JsFiles="" : ErrInfo="" : strCode="" : CurPage="" : strPage="" : i=0 : strSql="" : strMsg=""

Dim infoCode, infoTitle, infoContents, infoPreUrl, infoNewUrl			Rem 提示信息相关变量
infoCode="" : infoTitle="" : infoContents="" : infoPreUrl="" : infoNewUrl=""

IsDebug=True
%>
<!--#include file="const.asp"-->
<!--#include file="functions.asp"-->
<!--#include file="webclass.asp"-->
<!--#include file="styles.asp"-->
<!--#include file="mainmenu.asp"-->
<!--#include file="pagelink.asp"-->
<%
Call FileInc(1, "styles/styles.css")
Call FileInc(0, "js/jsfunc.js")
Call FileInc(0, "js/mouse_on_title.js")

Dim xjweb, Conn, Rs, ConnStr,dbPath
Set xjweb=new CLSXJWEB

dbPath="database/#2012jsb.mdb"
Connstr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(dbPath)
'ConnStr="sjjminfoman"	'运用系统DSN		用DSN联接时进行库中字段转换时会出问题

If IsNull(Session("userName")) Then Session("userName")=""
If IsNull(Session("userAble")) Or Session("userAble")="" Then Session("userAble")="000000000000000"
If IsNull(Session("userNick")) Then Session("userNick")=""

Call CheckCookies()
%>