<!--#include file="include/conn.asp"-->
<%
	'2016-01-06 16:49
	'���ļ�ֻ������Ӻ͸�����������������a
	Dim action
	action=Request("action")
	'���������б�����ʼ��	
	Dim strlsh, strddh, strmh, strdwmc, strdmmc, strtslb, sngmjzf, sngmtbl, sngmtjgbl, sngdxjgbl, strbz, strjgzz, strsjzz, dtjhkssj, dtjgjssj, dtjhjssj
	Dim strmjxx, strrwlr, dtrwxdsj, dtlzrq, strlzr, ibomzf, itsdzf, itszf, itsxxzlzf
	strlsh=Trim(UCase(Request("lsh"))) : strddh=Trim(Request("ddh")) : strdwmc=Trim(Request("dwmc")) : strdmmc=Trim(Request("dmmc"))
	strmh=Trim(Request("mh")) : strtslb=Request("tslb") : sngmjzf=Request("mjzf") : sngmtbl=Request("mtbl") : sngmtjgbl=Request("mtjgbl")
	sngdxjgbl=Request("dxjgbl") : strbz=Request("bz") : strjgzz=Request("jgzz") : strsjzz=Request("sjzz")
	dtjhkssj=Request("jhkssj") : dtjhjssj=Request("jhjssj") : dtjgjssj=Request("jgjssj") :  strmjxx=Request("mjxx") : strrwlr=Request("rwlr")
	dtrwxdsj=now() : dtlzrq=now() : strlzr=session("userName")
	ibomzf=Request("bomzf") : itsdzf=Request("tsdzf") : itszf=Request("tszf") : itsxxzlzf=Request("tsxxzlzf")

	sngmjzf=NullToNum(sngmjzf)
	sngmtbl=NulltoNum(sngmtbl)
	sngmtjgbl=NulltoNum(sngmtjgbl)
	sngdxjgbl=NulltoNum(sngdxjgbl)
	ibomzf=NulltoNum(ibomzf)
	itsdzf=NulltoNum(itsdzf)
	itszf=NulltoNum(itszf)
	itsxxzlzf=NulltoNum(itsxxzlzf)

	'�Դ�������ݽ��д���
	strMsg=""
	If strlsh="" Then strMsg="����С��Ϊ��!<br>"
	If strddh="" Then strMsg=strMsg & "������Ϊ��!<br>"
	If strdwmc="" Then strMsg=strMsg & "�ͻ�����Ϊ��!<br>"
	If strmh=""  Then strMsg=strMsg & "ԭ��ˮ��Ϊ��!<br>"
	If strtslb=""  Then strMsg=strMsg & "���������Ϊ��!<br>"
	If sngmjzf=0 Then strMsg=strMsg & "ģ���ܷ�Ϊ��!<br>"
	If sngmtjgbl=0 Then strMsg=strMsg & "ģͷ�ṹ����Ϊ��!<br>"
	If sngdxjgbl=0 Then strMsg=strMsg & "���ͽṹ����Ϊ��!<br>"
	If strjgzz="" or strsjzz="" Then strMsg=strMsg & "�鳤û��ѡ��!<br>"

	If strMsg <> "" and action <> "BzChan" Then
		infoTitle="���ݲ�����"
		infoContents=strMsg & "<br>���<a href=""#"" onclick='history.go(-1);'>����ǰҳ</a>��������"
		GotoPrompt()
	End If

	Call mtask_add()

	'������������
	Function mtask_add()
		'�����ˮ���Ƿ��Ѵ���
		Dim TmpRs
		Set TmpRs=xjweb.exec("select * from [mtask] where [lsh]='"&strlsh&"'",1)
		If Not(TmpRs.eof Or TmpRs.bof) Then
			If TmpRs("rwlr")<>"����" Then 
				Call JsAlert("��ˮ�� " & strlsh & " Ϊ"& TmpRs("rwlr") &"�����������ˮ��!","")
			else if isnull(TmpRs("mttsjs")) and isnull(TmpRs("dxtsjs")) Then
					strSql="select * from [mtask] where [lsh]='"& strlsh &"'"
					Call xjweb.exec("",-1)
					Rs.open strSql,Conn,1,3
					Rs("ddh")=strddh
					Rs("lsh")=strlsh
					Rs("dwmc")=strdwmc
					Rs("dmmc")=strdmmc
					Rs("mh")=strmh
					Rs("mjxx")=strmjxx
					Rs("rwlr")=strrwlr			
					If strtslb<>"" Then Rs("tslb")=strtslb
					If strbz<>"" Then Rs("bz")=strbz
					Rs("rwxdsj")=dtrwxdsj
					Rs("jhkssj")=dtjhkssj
					Rs("jhjgsj")=dtjgjssj
					Rs("jhjssj")=dtjhjssj
					Rs("jgzz")=strjgzz
					Rs("sjzz")=strsjzz
					Rs("lzr")=strlzr
					Rs("lzrq")=dtlzrq
					Rs("mjzf")=sngmjzf
					Rs("mtbl")=sngmtbl
					Rs("mtjgbl")=sngmtjgbl
					Rs("dxjgbl")=sngdxjgbl
					Rs("bomzf")=ibomzf
					Rs("tsdzf")=itsdzf
					Rs("tszf")=itszf
					Rs("tsxxzlzf")=itsxxzlzf
					Rs.update
					Rs.Close
					Call JsAlert("��ˮ�� " & strlsh &  " ��������ĳɹ�!", "Repair_add.asp")	
				else
					Call JsAlert("��ˮ�� " & strlsh & " ������������ɣ��޷�����������!","")
				End If			
			End If
		else
			strSql="select * from [mtask]"
			Call xjweb.exec("",-1)
			Rs.open strSql,Conn,1,3
			Rs.AddNew
				Rs("ddh")=strddh
				Rs("lsh")=strlsh
				Rs("dwmc")=strdwmc
				Rs("dmmc")=strdmmc
				Rs("mh")=strmh
				Rs("mjxx")=strmjxx
				Rs("rwlr")=strrwlr			
				If strtslb<>"" Then Rs("tslb")=strtslb
				If strbz<>"" Then Rs("bz")=strbz
				Rs("rwxdsj")=dtrwxdsj
				Rs("jhkssj")=dtjhkssj
				Rs("jhjgsj")=dtjgjssj
				Rs("jhjssj")=dtjhjssj
				Rs("jgzz")=strjgzz
				Rs("sjzz")=strsjzz
				Rs("lzr")=strlzr
				Rs("lzrq")=dtlzrq
				Rs("mjzf")=sngmjzf
				Rs("mtbl")=sngmtbl
				Rs("mtjgbl")=sngmtjgbl
				Rs("dxjgbl")=sngdxjgbl
				Rs("bomzf")=ibomzf
				Rs("tsdzf")=itsdzf
				Rs("tszf")=itszf
				Rs("tsxxzlzf")=itsxxzlzf
			Rs.update
			Rs.Close
			Call JsAlert("��������ӳɹ�!", "Repair_add.asp")
		End If
		TmpRs.Close		
	End Function
%>
