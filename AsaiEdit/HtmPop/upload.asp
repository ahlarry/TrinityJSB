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
'=======以下对上传文件后，返回不同信息作相应的提示===========
if request2.save("upfile",0) then
'	upload_Results="<SCRIPT language=javascript>" & vbcrlf	
'	upload_Results=upload_Results & "parent.document.pic.a.value='"&request2.form("upfile") &"';"& vbcrlf	& "</script>" 
	StrExt=Mid(request2.form("upfile"),InstrRev(request2.form("upfile"),".")+1)
	upload_Results="<SCRIPT language=javascript>" & vbcrlf
	upload_Results=upload_Results & "parent.document.pic.a.value='" &request2.form("upfile")& "';" & vbcrlf	
	upload_Results=upload_Results & "parent.document.pic.b.value='" &StrPath&request2.form("upfile")& "';" & vbcrlf	
	upload_Results=upload_Results & "parent.document.pic.c.value='" &StrExt& "';" & vbcrlf	
	upload_Results=upload_Results & "</script>"
	upload_Results=upload_Results & "附件：<span style=""color:#FF6633;"">"&request2.form("upfile")&"</span>"
else
	Err_num=request2.form("upfile_Err")
	upload_Results="上传失败！"
	if Err_num=-1 then
		upload_Results=upload_Results&"没有文件上传，请选择你要上传的文件"
	elseif Err_num=1 then
		upload_Results=upload_Results&"上传文件已超出最大限制(请小于"&request2.maxsize/1048576&"MB)"
	elseif Err_num=2 then
		upload_Results=upload_Results&"不允许上传此格式的文件"
	elseif Err_num=3 then
		upload_Results=upload_Results&"不允许上传此格式的文件，而且上传文件已超过最大限制"
	elseif Err_num=4 then
		upload_Results=upload_Results&"文件含有恶意代码，不允许上传"
	else
		upload_Results=upload_Results&"上传过程中发生意外"
	end if
	upload_Results="<SCRIPT language=javascript>" & vbcrlf & "alert('" & upload_Results & "');" & vbcrlf & "</script>"
	upload_Results=upload_Results & " <a href='UpForm.asp'>返回,重新上传？</a>"
end if
response.write upload_Results
set request2=nothing
%>
</body>
</html>
