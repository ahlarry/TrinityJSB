<%
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
Folders=split(Folders,"+") '转化为数组
keyword=trim(keyword) '去掉前后空格
if keyword="" then
Response.Write("<font color='red'>关键字不能为空</font><br/>")
exit Function
end if
'判断是否包含非法字符
flag=instr(keyword,"\") or instr(keyword,"/")
flag=flag or instr(keyword,":")
flag=flag or instr(keyword,"|")
flag=flag or instr(keyword,"&")
if flag then '关键字中不能包含\/:|&
Response.Write("<font color='red'>关键字不能包含/\:|&</font><br/>")
Exit Function '如果包含有这个则退出
end if
'多路径搜索
dim i
for i=0 to ubound(Folders)
Call GetAllFile(Folders(i)) '调用循环递归函数
next
Response.Write("共搜索到<font color='red'>"&Counter&"</font>个结果")
End Function
'***************历遍文件和文件夹******************************
Private Function GetAllFile(Folder)
dim objFd,objFs,objFf
Set objFd=objFso.GetFolder(Folder)
Set objFs=objFd.SubFolders
Set objFf=objFd.Files
'历遍子文件夹
dim strFdName '声明子文件夹名
'*********历遍子文件夹******
on error resume next
For Each OneDir In objFs
strFdName=OneDir.Name
'系统文件夹不在历遍之列
If strFdName<>"Config.Msi" EQV strFdName<>"RECYCLED" EQV strFdName<>"RECYCLER" EQV strFdName<>"System Volume Information" Then
SFN=Folder&"\"&strFdName '绝对路径
Call GetAllFile(SFN) '调用递归
End If
Next
dim strFlName
'**********历遍文件********
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
Dim objReg
objReg=Right(FileName,(Len(FileName)-InStrRev(FileName,"\")))
If STRCOMP(objReg,keyword,1)=0 Then
FileName="\\sj901"& Mid(FileName,InStr(FileName,"\"))
OutPut="<a href="&FileName&">"&FileName&"</a><br/>"
'OutPut="<a href='#'>"&FileName&"</a><br/>"
'Response.Write(objFd&"<p>")
Response.Write(OutPut&"<p>") '输出匹配的结果
''**************************搜索并使关键字上色*************************
'Private Function ColorOn(FileName)
'dim objReg
'Set objReg=new RegExp
'objReg.Pattern=CreatePattern(keyword)
'objReg.IgnoreCase=True
'objReg.Global=True
'retVal=objReg.Test(FileName) '进行搜索测试,如果通过则上色并输出
'if retVal then
'OutPut=objReg.Replace(FileName,"<font color='#FF0000'>$1</font>") '设置关键字的显示颜色
''***************************该部分可以根据需要修改输出************************************
'OutPut="<a href='#'>"&OutPut&"</a><br/>"
'Response.Write(OutPut&"<p>") '输出匹配的结果
'Response.Write(keyword&"<p>")
'Response.Write(FileName&"☆<p>")
'*************************************可修改部分结束**************************************
ColorOn=1 '加入计数器的数目
else
ColorOn=0
end if
Set objReg=Nothing
End Function
End Class
'************************结束类SearchFile**********************
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>模具断面图搜索</title>
</head>
<body>
<form name="form1" method="post" action="<% =Request.ServerVariables("PATH_INFO")%>">
请输入模具流水号:
<input name="keyword" type="text" id="keyword">
<input type="submit" name="Submit" value="搜索">
</form>
<%
Dim s_lsh, action
s_lsh=Trim(Request("s_lsh"))
If s_lsh<>"" Then
	keyword=s_lsh &".dwg"
Else
	keyword=Request.Form("keyword") &".dwg"
End If
if keyword<>"" then
Set newsearch=new SearchFile
'newsearch.Folders=Server.mappath("dmtj")
newsearch.Folders="G:\设计参考\断面图集" '是绝对路径
newsearch.keyword=keyword
newsearch.Search
Set newsearch=Nothing
response.Write("<br/>费时："&(timer()-st)*1000&"毫秒")
end if
%>
</body>
</html>
