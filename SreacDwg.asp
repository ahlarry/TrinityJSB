<%
'*************Set newsearch=new SearchFile '���� *************
'*************newsearch.Folder="F:+E:"'��������Դ*************
'*************newsearch.keyword="���" '�ؼ���*************
'*************newsearch.Search '��ʼ����*************
'*************Set newsearch=Nothing '����*************
'*************************************************************
Server.ScriptTimeOut =99999 '������صĳ�ʱ����
Class SearchFile
dim Folders '�������·��,��·��ʹ��+������,�����пո�
dim keyword '����ؼ���
dim objFso '����ȫ�ֱ���
dim Counter '����ȫ�ֱ����������������Ŀ
'*****************��ʼ��**************************************
Private Sub Class_Initialize
Set objFso=Server.CreateObject("Scripting.FileSystemObject")
Counter=0 '��ʼ��������
End Sub
'************************************************************
Private Sub Class_Terminate
Set objFso=Nothing
End Sub
'**************���г�Ա,���õķ���***************************
Function Search
Folders=split(Folders,"+") 'ת��Ϊ����
keyword=trim(keyword) 'ȥ��ǰ��ո�
if keyword="" then
Response.Write("<font color='red'>�ؼ��ֲ���Ϊ��</font><br/>")
exit Function
end if
'�ж��Ƿ�����Ƿ��ַ�
flag=instr(keyword,"\") or instr(keyword,"/")
flag=flag or instr(keyword,":")
flag=flag or instr(keyword,"|")
flag=flag or instr(keyword,"&")
if flag then '�ؼ����в��ܰ���\/:|&
Response.Write("<font color='red'>�ؼ��ֲ��ܰ���/\:|&</font><br/>")
Exit Function '���������������˳�
end if
'��·������
dim i
for i=0 to ubound(Folders)
Call GetAllFile(Folders(i)) '����ѭ���ݹ麯��
next
Response.Write("��������<font color='red'>"&Counter&"</font>�����")
End Function
'***************�����ļ����ļ���******************************
Private Function GetAllFile(Folder)
dim objFd,objFs,objFf
Set objFd=objFso.GetFolder(Folder)
Set objFs=objFd.SubFolders
Set objFf=objFd.Files
'�������ļ���
dim strFdName '�������ļ�����
'*********�������ļ���******
on error resume next
For Each OneDir In objFs
strFdName=OneDir.Name
'ϵͳ�ļ��в�������֮��
If strFdName<>"Config.Msi" EQV strFdName<>"RECYCLED" EQV strFdName<>"RECYCLER" EQV strFdName<>"System Volume Information" Then
SFN=Folder&"\"&strFdName '����·��
Call GetAllFile(SFN) '���õݹ�
End If
Next
dim strFlName
'**********�����ļ�********
For Each OneFile In objFf
strFlName=OneFile.Name
'desktop.ini��folder.htt���ص�ϵͳ�ļ�������ȡ��Χ
If strFlName<>"desktop.ini" EQV strFlName<>"folder.htt" Then
FN=Folder&"\"&strFlName
Counter=Counter+ColorOn(FN)
End If
Next
'***************************
'�رո�����ʵ��
Set objFd=Nothing
Set objFs=Nothing
Set objFf=Nothing
End Function
'*********************����ƥ��ģʽ***********************************
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
CreatePattern=Replace(CreatePattern,"*","[^\\\/]*") '*��ƥ��
CreatePattern=Replace(CreatePattern,"?","[^\\\/]{1}") '?��ƥ��
CreatePattern="("&CreatePattern&")+" '����ƥ��
End Function
'***************************������ȫƥ����ļ���*********************
Private Function ColorOn(FileName)
Dim objReg
objReg=Right(FileName,(Len(FileName)-InStrRev(FileName,"\")))
If STRCOMP(objReg,keyword,1)=0 Then
FileName="\\sj901"& Mid(FileName,InStr(FileName,"\"))
OutPut="<a href="&FileName&">"&FileName&"</a><br/>"
'OutPut="<a href='#'>"&FileName&"</a><br/>"
'Response.Write(objFd&"<p>")
Response.Write(OutPut&"<p>") '���ƥ��Ľ��
''**************************������ʹ�ؼ�����ɫ*************************
'Private Function ColorOn(FileName)
'dim objReg
'Set objReg=new RegExp
'objReg.Pattern=CreatePattern(keyword)
'objReg.IgnoreCase=True
'objReg.Global=True
'retVal=objReg.Test(FileName) '������������,���ͨ������ɫ�����
'if retVal then
'OutPut=objReg.Replace(FileName,"<font color='#FF0000'>$1</font>") '���ùؼ��ֵ���ʾ��ɫ
''***************************�ò��ֿ��Ը�����Ҫ�޸����************************************
'OutPut="<a href='#'>"&OutPut&"</a><br/>"
'Response.Write(OutPut&"<p>") '���ƥ��Ľ��
'Response.Write(keyword&"<p>")
'Response.Write(FileName&"��<p>")
'*************************************���޸Ĳ��ֽ���**************************************
ColorOn=1 '�������������Ŀ
else
ColorOn=0
end if
Set objReg=Nothing
End Function
End Class
'************************������SearchFile**********************
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>ģ�߶���ͼ����</title>
</head>
<body>
<form name="form1" method="post" action="<% =Request.ServerVariables("PATH_INFO")%>">
������ģ����ˮ��:
<input name="keyword" type="text" id="keyword">
<input type="submit" name="Submit" value="����">
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
newsearch.Folders="G:\��Ʋο�\����ͼ��" '�Ǿ���·��
newsearch.keyword=keyword
newsearch.Search
Set newsearch=Nothing
response.Write("<br/>��ʱ��"&(timer()-st)*1000&"����")
end if
%>
</body>
</html>
