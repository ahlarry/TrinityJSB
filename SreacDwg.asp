<!--#include file="include/conn.asp"-->
<%
CurPage="断面图和投影图"					'页面的名称位置( 任务书管理 → 添加任务书)
strPage=""
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, strlsh
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder
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
newsearch.Folders="G:\设计参考\断面图集+F:\模具图档\模具修理" '是绝对路径
'newsearch.Folders="D:\模具图\+G:\迅雷下载" '是绝对路径
newsearch.keyword=keyword
newsearch.Search
Set newsearch=Nothing
end if
Call BottomTable()
xjweb.footer()
closeObj()
'*************Set newsearch=new SearchFile '声明 *************
'*************newsearch.Folder="F:+E:"'传入搜索源*************
'*************newsearch.keyword="汇编" '关键词*************
'*************newsearch.Search '开始搜索*************
'*************Set newsearch=Nothing '结束*************
'*************************************************************
Server.ScriptTimeOut =99999 '程序加载的超时设置
Class SearchFile
dim Folders '传入绝对路径,多路径使用+号连接,不能有空格
dim keyword '传入关键词
dim objFso '定义全局变量
dim Counter '定义全局变量，搜索结果的数目
'*****************初始化**************************************
Private Sub Class_Initialize
Set objFso=Server.CreateObject("Scripting.FileSystemObject")
Counter=0 '初始化计数器
End Sub
'************************************************************
Private Sub Class_Terminate
Set objFso=Nothing
End Sub
'**************公有成员,调用的方法***************************
Function Search
dim flag
Folders=split(Folders,"+") '转化为数组
keyword=trim(keyword) '去掉前后空格
if keyword="" then
Call JsAlert("关键字不能为空","")
exit Function
end if
'判断是否包含非法字符
flag=instr(keyword,"\") or instr(keyword,"/")
flag=flag or instr(keyword,":")
flag=flag or instr(keyword,"|")
flag=flag or instr(keyword,"&")
if flag then '关键字中不能包含\/:|&
Call JsAlert("关键字不能包含/\:|&","")
Exit Function '如果包含有这个则退出
end if
'多路径搜索
dim i
for i=0 to ubound(Folders)
Call GetAllFile(Folders(i)) '调用循环递归函数
next
End Function
'***************历遍文件和文件夹******************************
Private Function GetAllFile(Folder)
dim objFd,objFs,objFf
Set objFd=objFso.GetFolder(Folder)
Set objFs=objFd.SubFolders
Set objFf=objFd.Files
'*********历遍子文件夹******
dim OneDir,SFN,strFdName  '声明子文件夹名
on error resume next
For Each OneDir In objFs
strFdName=OneDir.Name
'系统文件夹不在历遍之列
If strFdName<>"Config.Msi" EQV strFdName<>"RECYCLED" EQV strFdName<>"RECYCLER" EQV strFdName<>"System Volume Information" Then
SFN=Folder&"\"&strFdName '绝对路径
Call GetAllFile(SFN) '调用递归
End If
Next
'**********历遍文件********
dim strFlName,OneFile,FN
For Each OneFile In objFf
strFlName=OneFile.Name
'desktop.ini和folder.htt隐藏的系统文件不在列取范围
If strFlName<>"desktop.ini" EQV strFlName<>"folder.htt" Then
FN=Folder&"\"&strFlName
Counter=Counter+ColorOn(FN)
End If
Next
'***************************
'关闭各对象实例
Set objFd=Nothing
Set objFs=Nothing
Set objFf=Nothing
End Function
'*********************生成匹配模式***********************************
Private Function CreatePattern(keyword)
CreatePattern=keyword
CreatePattern=Replace(CreatePattern,".","\.")
CreatePattern=Replace(CreatePattern,"+","\+")
CreatePattern=Replace(CreatePattern,"(","\(")
CreatePattern=Replace(CreatePattern,")","\)")
CreatePattern=Replace(CreatePattern,"[","\[")
CreatePattern=Replace(CreatePattern,"]","\]")
CreatePattern=Replace(CreatePattern,"{","\{")
CreatePattern=Replace(CreatePattern,"}","\}")
CreatePattern=Replace(CreatePattern,"*","[^\\\/]*") '*号匹配
CreatePattern=Replace(CreatePattern,"?","[^\\\/]{1}") '?号匹配
CreatePattern="("&CreatePattern&")+" '整体匹配
End Function
'***************************搜索完全匹配的文件名*********************
Private Function ColorOn(FileName)
Dim objReg,objJpg,OutPut
'Response.Write("---"&FileName&"<p>")'=======调试输出======
objReg=Right(FileName,(Len(FileName)-InStrRev(FileName,"\")))
objJpg=Right(FileName,(Len(FileName)-InStrRev(FileName,".")))
If STRCOMP(objReg,keyword&".dwg",1)=0 Then
FileName="\\sj901"& Mid(FileName,InStr(FileName,"\"))
OutPut="<a href="&FileName&">"&FileName&"</a><br/>"
Response.Write("<p>"&OutPut&"<p>") '输出匹配的结果
ColorOn=1 '加入计数器的数目
else if objJpg="jpg" and Instr(FileName,keyword)>0 Then
FileName="\\sj901"& Mid(FileName,InStr(FileName,"\"))
'OutPut="<a href="&FileName&">"&FileName&"</a><br/>"
OutPut="<img src=file:\\"&FileName&"  onload='javascript:if(this.width>("&web_info(8)&"-10)) this.width=("&web_info(8)&"-10);' border='0'></img>"
Response.Write("<p>"&OutPut&"<p>"&objReg&"<p>==========<p>") '输出匹配的结果
ColorOn=1 '加入计数器的数目
End If
ColorOn=0
end if
Set objReg=Nothing
Set objJpg=Nothing
End Function
End Class
'************************结束类SearchFile**********************
%>