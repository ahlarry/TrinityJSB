<!--#include file="include/conn.asp"-->
<%
CurPage="����ͼ��ͶӰͼ"					'ҳ�������λ��( ��������� �� ���������)
strPage=""
xjweb.header()
Call TopTable()
Dim strFeedBack, strOrder, strO, strlsh
strOrder=Trim(Request("order"))
strFeedBack="&order="&strOrder
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
newsearch.Folders="G:\��Ʋο�\����ͼ��+F:\ģ��ͼ��\ģ������" '�Ǿ���·��
'newsearch.Folders="D:\ģ��ͼ\+G:\Ѹ������" '�Ǿ���·��
newsearch.keyword=keyword
newsearch.Search
Set newsearch=Nothing
end if
Call BottomTable()
xjweb.footer()
closeObj()
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
dim flag
Folders=split(Folders,"+") 'ת��Ϊ����
keyword=trim(keyword) 'ȥ��ǰ��ո�
if keyword="" then
Call JsAlert("�ؼ��ֲ���Ϊ��","")
exit Function
end if
'�ж��Ƿ�����Ƿ��ַ�
flag=instr(keyword,"\") or instr(keyword,"/")
flag=flag or instr(keyword,":")
flag=flag or instr(keyword,"|")
flag=flag or instr(keyword,"&")
if flag then '�ؼ����в��ܰ���\/:|&
Call JsAlert("�ؼ��ֲ��ܰ���/\:|&","")
Exit Function '���������������˳�
end if
'��·������
dim i
for i=0 to ubound(Folders)
Call GetAllFile(Folders(i)) '����ѭ���ݹ麯��
next
End Function
'***************�����ļ����ļ���******************************
Private Function GetAllFile(Folder)
dim objFd,objFs,objFf
Set objFd=objFso.GetFolder(Folder)
Set objFs=objFd.SubFolders
Set objFf=objFd.Files
'*********�������ļ���******
dim OneDir,SFN,strFdName  '�������ļ�����
on error resume next
For Each OneDir In objFs
strFdName=OneDir.Name
'ϵͳ�ļ��в�������֮��
If strFdName<>"Config.Msi" EQV strFdName<>"RECYCLED" EQV strFdName<>"RECYCLER" EQV strFdName<>"System Volume Information" Then
SFN=Folder&"\"&strFdName '����·��
Call GetAllFile(SFN) '���õݹ�
End If
Next
'**********�����ļ�********
dim strFlName,OneFile,FN
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
Dim objReg,objJpg,OutPut
'Response.Write("---"&FileName&"<p>")'=======�������======
objReg=Right(FileName,(Len(FileName)-InStrRev(FileName,"\")))
objJpg=Right(FileName,(Len(FileName)-InStrRev(FileName,".")))
If STRCOMP(objReg,keyword&".dwg",1)=0 Then
FileName="\\sj901"& Mid(FileName,InStr(FileName,"\"))
OutPut="<a href="&FileName&">"&FileName&"</a><br/>"
Response.Write("<p>"&OutPut&"<p>") '���ƥ��Ľ��
ColorOn=1 '�������������Ŀ
else if objJpg="jpg" and Instr(FileName,keyword)>0 Then
FileName="\\sj901"& Mid(FileName,InStr(FileName,"\"))
'OutPut="<a href="&FileName&">"&FileName&"</a><br/>"
OutPut="<img src=file:\\"&FileName&"  onload='javascript:if(this.width>("&web_info(8)&"-10)) this.width=("&web_info(8)&"-10);' border='0'></img>"
Response.Write("<p>"&OutPut&"<p>"&objReg&"<p>==========<p>") '���ƥ��Ľ��
ColorOn=1 '�������������Ŀ
End If
ColorOn=0
end if
Set objReg=Nothing
Set objJpg=Nothing
End Function
End Class
'************************������SearchFile**********************
%>