<!--#include file="include/conn.asp"-->
<%
	Call ChkAble(5)
	'���ļ�ֻ���鳤��������������
	dim strmtjgr, strdxjgr, strgjjgr, strmtsjr, strdxsjr, strgjsjr, strmtshr, strdxshr, strgjshr, strmtbomr, strdxbomr, strlsh,strmtjgshr, strmtsjshr, strdxjgshr, strdxsjshr, strgjjgshr, strgjsjshr, strmtjgsj, strmtsjsj, strmtshsj, strmtbomsj, strmtjgshsj, strmtsjshsj, strdxjgsj, strdxsjsj, strdxshsj, strdxbomsj, strgjshsj, strgjjgshsj, strgjsjshsj, strdxjgshsj, strdxsjshsj, strgjjgsj, strgjsjsj, strgjfcr, strgjfcsj, strmtgysjr, strmtgysjjs, strdxgysjr, strdxgysjjs, strmtgyshr, strmtgyshjs, strdxgyshr, strdxgyshjs, strgjgysjr, strgjgysjks, strgjgysjjs, strgjgyshr, strgjgyshks, strgjgyshjs
	strmtjgr="" : strdxjgr="" : strgjjgr="" : strmtsjr="" : strdxsjr="" : strgjsjr=""
	strmtshr="" : strdxshr="" : strgjshr="" : strmtbomr="" : strdxbomr="" : strgjfcr="" : strgjfcsj=""
	strmtjgshr="" : strmtsjshr="" : strdxjgshr="" : strdxsjshr="" : strgjjgshr="" : strgjsjshr=""
	strmtgysjr="" : strdxgysjr="" : strmtgyshr="" : strdxgyshr="" :	strgjgysjr="" : strgjgyshr=""


	strmtjgsj=Replace(request("mtjgsj"),"."," ")
	strmtsjsj=Replace(request("mtsjsj"),"."," ")
	strmtshsj=Replace(request("mtshsj"),"."," ")
	strmtjgshsj=Replace(request("mtjgshsj"),"."," ")
	strmtsjshsj=Replace(request("mtsjshsj"),"."," ")
	strdxjgsj=Replace(request("dxjgsj"),"."," ")
	strdxsjsj=Replace(request("dxsjsj"),"."," ")
	strdxshsj=Replace(request("dxshsj"),"."," ")
	strmtbomsj=Replace(request("mtbomsj"),"."," ")
	strdxbomsj=Replace(request("mtbomsj"),"."," ")
	strdxjgshsj=Replace(request("dxjgshsj"),"."," ")
	strdxsjshsj=Replace(request("dxsjshsj"),"."," ")
	strgjjgsj=Replace(request("gjjgsj"),"."," ")
	strgjsjsj=Replace(request("gjsjsj"),"."," ")
	strgjshsj=Replace(request("gjshsj"),"."," ")
	strgjjgshsj=Replace(request("gjjgshsj"),"."," ")
	strgjsjshsj=Replace(request("gjsjshsj"),"."," ")
	strgjfcsj=Replace(request("gjfcsj"),"."," ")

	strmtgysjjs=Replace(request("mtgysjsj"),"."," ")
	strdxgysjjs=Replace(request("dxgysjsj"),"."," ")
	strgjgysjjs=Replace(request("gjgysjsj"),"."," ")
	strmtgyshjs=Replace(request("mtgyshsj"),"."," ")
	strdxgyshjs=Replace(request("dxgyshsj"),"."," ")
	strgjgyshjs=Replace(request("gjgyshsj"),"."," ")

	strmtjgr=request("mtjgr")
	strmtsjr=request("mtsjr")
	strmtshr=request("mtshr")
	strmtbomr=request("mtbomr")
	strdxjgr=request("dxjgr")
	strdxsjr=request("dxsjr")
	strdxshr=request("dxshr")
	strdxbomr=request("dxbomr")
	strgjjgr=request("gjjgr")
	strgjsjr=request("gjsjr")
	strgjshr=request("gjshr")
	strlsh=request("lsh")
	strmtjgshr=request("mtjgshr")
	strmtsjshr=request("mtsjshr")
	strdxjgshr=request("dxjgshr")
	strdxsjshr=request("dxsjshr")
	strgjjgshr=request("gjjgshr")
	strgjsjshr=request("gjsjshr")
	strgjfcr=request("gjfcr")

	strmtgysjr=request("mtgysjr")
	strdxgysjr=request("dxgysjr")
	strmtgyshr=request("mtgyshr")
	strdxgyshr=request("dxgyshr")
	strgjgysjr=request("gjgysjr")
	strgjgyshr=request("gjgyshr")

	If strlsh="" Then
		Call JsAlert("�޷�ȷ�����������ˮ��,���������ڽ���!", "")
	End If

	Set Rs=xjweb.Exec("select lsh from [mtask] where [lsh]='"&strlsh&"'",1)
	If Rs.Eof Or Rs.Bof Then
		Call JsAlert("��ˮ�� ��" & strlsh &  "�� �����鲻����!","")
	End If
	Rs.Close

'		Call JsAlert(strmtbomsj,"")

	strSql="select * from [mtask] where [lsh]='" & strlsh & "'"
	Call xjweb.Exec("",-1)
	strMsg=""
	Rs.open strsql,Conn,1,3
		if strmtjgr<>"" and strmtjgr<>rs("mtjgr") then rs("mtjgr")=strmtjgr : strmsg = strmsg & "����ģͷ�ṹ��"
		if strmtjgsj<>"" and strmtjgr<>rs("mtjgjs") then rs("mtjgjs")=strmtjgsj : strmsg = strmsg & "����ģͷ�ṹ����ʱ��"
		if strdxjgr<>"" and strdxjgr<>rs("dxjgr") then rs("dxjgr")=strdxjgr : strmsg = strmsg & "���Ķ��ͽṹ��"
		if strdxjgsj<>"" and strdxjgsj<>rs("dxjgjs") then rs("dxjgjs")=strdxjgsj : strmsg = strmsg & "���Ķ��ͽṹ����ʱ��"
		if strgjjgr<>"" and strgjjgr<>rs("gjjgr") then rs("gjjgr")=strgjjgr : strmsg = strmsg & "���ĺ󹲼��ṹ��"
		if strgjjgsj<>"" and strgjjgsj<>rs("gjjgjs") then rs("gjjgjs")=strgjjgsj : strmsg = strmsg & "���ĺ󹲼��ṹ����ʱ��"
		if strmtsjr<>"" and strmtsjr<>rs("mtsjr") then rs("mtsjr")=strmtsjr : strmsg = strmsg & "����ģͷ�����"
		if strmtsjsj<>"" and strmtsjsj<>rs("mtsjjs") then rs("mtsjjs")=strmtsjsj : strmsg = strmsg & "����ģͷ��ƽ���ʱ��"
		if strdxsjr<>"" and strdxsjr<>rs("dxsjr") then rs("dxsjr")=strdxsjr : strmsg = strmsg & "���Ķ��������"
		if strdxsjsj<>"" and strdxsjsj<>rs("dxsjjs") then rs("dxsjjs")=strdxsjsj : strmsg = strmsg & "���Ķ�����ƽ���ʱ��"
		if strgjsjr<>"" and strgjsjr<>rs("gjsjr") then rs("gjsjr")=strgjsjr : strmsg = strmsg & "���ĺ󹲼������"
		if strgjsjsj<>"" and strgjsjsj<>rs("gjsjjs") then rs("gjsjjs")=strgjsjsj : strmsg = strmsg & "���ĺ󹲼���ƽ���ʱ��"
		if strmtshr<>"" and strmtshr<>rs("mtshr") then rs("mtshr")=strmtshr : strmsg = strmsg & "����ģͷ�����"
		if strmtshsj<>"" and strmtshsj<>rs("mtshjs") then rs("mtshjs")=strmtshsj : strmsg = strmsg & "����ģͷ��˽���ʱ��"
		if strdxshr<>"" and strdxshr<>rs("dxshr") then rs("dxshr")=strdxshr : strmsg = strmsg & "���Ķ��������"
		if strdxshsj<>"" and strdxshsj<>rs("dxshjs") then rs("dxshjs")=strdxshsj : strmsg = strmsg & "���Ķ�����˽���ʱ��"
		if strgjshr<>"" and strgjshr<>rs("gjshr") then rs("gjshr")=strgjshr : strmsg = strmsg & "���ĺ󹲼������"
		if strgjshsj<>"" and strgjshsj<>rs("gjshjs") then rs("gjshjs")=strgjshsj : strmsg = strmsg & "���ĺ󹲼���˽���ʱ��"
		if strmtbomr<>"" and strmtbomr<>rs("mtbomr") then rs("mtbomr")=strmtbomr : strmsg = strmsg & "����ģͷBOM��"
		if strmtbomsj<>"" and strmtbomsj<>rs("mtbomjs") then rs("mtbomjs")=strmtbomsj : strmsg = strmsg & "����ģͷBOM����ʱ��"
		if strdxbomr<>"" and strdxbomr<>rs("dxbomr") then rs("dxbomr")=strdxbomr : strmsg = strmsg & "���Ķ���BOM��"
		if strdxbomsj<>"" and strdxbomsj<>rs("dxbomjs") then rs("dxbomjs")=strdxbomsj : strmsg = strmsg & "���Ķ���BOM����ʱ��"
		if strmtjgshr<>"" and strmtjgshr<>rs("mtjgshr") then rs("mtjgshr")=strmtjgshr : strmsg = strmsg & "����ģͷ�ṹȷ����"
		if strmtjgshsj<>"" and strmtjgshsj<>rs("mtjgshjs") then rs("mtjgshjs")=strmtjgshsj : strmsg = strmsg & "����ģͷ�ṹȷ�Ͻ���ʱ��"
		if strmtsjshr<>"" and strmtsjshr<>rs("mtsjshr") then rs("mtsjshr")=strmtsjshr : strmsg = strmsg & "����ģͷ���ȷ����"
		if strmtsjshsj<>"" and strmtsjshsj<>rs("mtsjshjs") then rs("mtsjshjs")=strmtsjshsj : strmsg = strmsg & "����ģͷ���ȷ�Ͻ���ʱ��"
		if strdxjgshr<>"" and strdxjgshr<>rs("dxjgshr") then rs("dxjgshr")=strdxjgshr : strmsg = strmsg & "���Ķ��ͽṹȷ����"
		if strdxjgshsj<>"" and strdxjgshsj<>rs("dxjgshjs") then rs("dxjgshjs")=strdxjgshsj : strmsg = strmsg & "���Ķ��ͽṹȷ�Ͻ���ʱ��"
		if strdxsjshr<>"" and strdxsjshr<>rs("dxsjshr") then rs("dxsjshr")=strdxsjshr : strmsg = strmsg & "���Ķ������ȷ����"
		if strdxsjshsj<>"" and strdxsjshsj<>rs("dxsjshjs") then rs("dxsjshjs")=strdxsjshsj : strmsg = strmsg & "���Ķ������ȷ�Ͻ���ʱ��"
		if strgjjgshr<>"" and strgjjgshr<>rs("gjjgshr") then rs("gjjgshr")=strgjjgshr : strmsg = strmsg & "���ĺ󹲼��ṹȷ����"
		if strgjjgshsj<>"" and strgjjgshsj<>rs("gjjgshjs") then rs("gjjgshjs")=strgjjgshsj : strmsg = strmsg & "���ĺ󹲼��ṹȷ�Ͻ���ʱ��"
		if strgjsjshr<>"" and strgjsjshr<>rs("gjsjshr") then rs("gjsjshr")=strgjsjshr : strmsg = strmsg & "���ĺ󹲼����ȷ����"
		if strgjsjshsj<>"" and strgjsjshsj<>rs("gjsjshjs") then rs("gjsjshjs")=strgjsjshsj : strmsg = strmsg & "���ĺ󹲼����ȷ�Ͻ���ʱ��"
		if strgjfcr<>"" and strgjfcr<>rs("gjshr") then rs("gjshr")=strgjfcr : strmsg = strmsg & "���Ĺ���������"
		if strgjfcsj<>"" and strgjfcsj<>rs("gjshjs") then rs("gjshjs")=strgjfcsj : strmsg = strmsg & "���Ĺ����������ʱ��"

		if strmtgysjr<>"" and strmtgysjr<>rs("mtgysjr") then rs("mtgysjr")=strmtgysjr : strmsg = strmsg & "����ģͷ���������"
		if strmtgysjjs<>"" and strmtgysjjs<>rs("mtgysjjs") then rs("mtgysjjs")=strmtgysjjs : strmsg = strmsg & "����ģͷ������ƽ���ʱ��"
		if strmtgyshr<>"" and strmtgyshr<>rs("mtgyshr") then rs("mtgyshr")=strmtgyshr : strmsg = strmsg & "����ģͷ���������"
		if strmtgyshjs<>"" and strmtgyshjs<>rs("mtgyshjs") then rs("mtgyshjs")=strmtgyshjs : strmsg = strmsg & "����ģͷ������˽���ʱ��"
		if strdxgysjr<>"" and strdxgysjr<>rs("dxgysjr") then rs("dxgysjr")=strdxgysjr : strmsg = strmsg & "���Ķ��͹��������"
		if strdxgysjjs<>"" and strdxgysjjs<>rs("dxgysjjs") then rs("dxgysjjs")=strdxgysjjs : strmsg = strmsg & "���Ķ��͹�����ƽ���ʱ��"
		if strdxgyshr<>"" and strdxgyshr<>rs("dxgyshr") then rs("dxgyshr")=strdxgyshr : strmsg = strmsg & "���Ķ��͹��������"
		if strdxgyshjs<>"" and strdxgyshjs<>rs("dxgyshjs") then rs("dxgyshjs")=strdxgyshjs : strmsg = strmsg & "���Ķ��͹�����˽���ʱ��"

		if strgjgysjr<>"" and strgjgysjr<>rs("gjgysjr") then rs("gjgysjr")=strgjgysjr : strmsg = strmsg & "���Ĺ������������"
		if strgjgysjjs<>"" and strgjgysjjs<>rs("gjgysjjs") then rs("gjgysjjs")=strgjgysjjs : strmsg = strmsg & "���Ĺ���������ƽ���ʱ��"
		if strgjgyshr<>"" and strgjgyshr<>rs("gjgyshr") then rs("gjgyshr")=strgjgyshr : strmsg = strmsg & "���Ĺ������������"
		if strgjgyshjs<>"" and strgjgyshjs<>rs("gjgyshjs") then rs("gjgyshjs")=strgjgyshjs : strmsg = strmsg & "���Ĺ���������˽���ʱ��"
	Rs.update
	Rs.Close

	If strMsg <> "" Then
		strMsg = "���ݿ����: " & strMsg
		strSql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','����������','"&strmsg&"','"&now()&"')"
		Call xjweb.Exec(strSql,0)
		Call JsAlert("��ˮ�� �� " & strlsh & " ��������������˸��ĳɹ�!", "mtask_zzchange.asp")
	Else
		Call JsAlert("��û�н����κθ���!","mtask_zzchange.asp")
	End If
%>