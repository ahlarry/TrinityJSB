<!--#include file="../UpLoadClass.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
</head>

<body>
<%
Dim upload_Results, StrExt
StrExt=""
upload_Results=""
On Error Resume Next
set request2=new UpLoadClass
request2.autosave=2
request2.SavePath="../FileUp/"&formatdatetime(now(),2)&"/"
request2.open()
StrPath="AsaiEdit/FileUp/"&formatdatetime(now(),2)&"/"
'=======���¶��ϴ��ļ��󣬷��ز�ͬ��Ϣ����Ӧ����ʾ===========
if request2.save("upfile",0) then
'	upload_Results="<SCRIPT language=javascript>" & vbcrlf	
'	upload_Results=upload_Results & "parent.document.pic.a.value='"&request2.form("upfile") &"';"& vbcrlf	& "</script>" 
	StrExt=Mid(request2.form("upfile"),InstrRev(request2.form("upfile"),".")+1)
	upload_Results="<SCRIPT language=javascript>" & vbcrlf
	upload_Results=upload_Results & "parent.document.pic.a.value='" &request2.form("upfile")& "';" & vbcrlf	
	upload_Results=upload_Results & "parent.document.pic.b.value='" &StrPath&request2.form("upfile")& "';" & vbcrlf	
	upload_Results=upload_Results & "parent.document.pic.c.value='" &StrExt& "';" & vbcrlf	
	upload_Results=upload_Results & "</script>"
	upload_Results=upload_Results & "������<span style=""color:#FF6633;"">"&request2.form("upfile")&"</span>"
else
	Err_num=request2.form("upfile_Err")
	upload_Results="�ϴ�ʧ�ܣ�"
	if Err_num=-1 then
		upload_Results=upload_Results&"û���ļ��ϴ�����ѡ����Ҫ�ϴ����ļ�"
	elseif Err_num=1 then
		upload_Results=upload_Results&"�ϴ��ļ��ѳ����������(��С��"&request2.maxsize/1048576&"MB)"
	elseif Err_num=2 then
		upload_Results=upload_Results&"�������ϴ��˸�ʽ���ļ�"
	elseif Err_num=3 then
		upload_Results=upload_Results&"�������ϴ��˸�ʽ���ļ��������ϴ��ļ��ѳ����������"
	elseif Err_num=4 then
		upload_Results=upload_Results&"�ļ����ж�����룬�������ϴ�"
	else
		upload_Results=upload_Results&"�ϴ������з�������"
	end if
	upload_Results="<SCRIPT language=javascript>" & vbcrlf & "alert('" & upload_Results & "');" & vbcrlf & "</script>"
	upload_Results=upload_Results & " <a href='UpForm.asp'>����,�����ϴ���</a>"
end if
response.write upload_Results
set request2=nothing
%>
</body>
</html>
