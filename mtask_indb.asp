<!--#include file="include/conn.asp"-->
<%
'14:22 2007-1-6-������
	'���ļ�ֻ������Ӻ͸�������������
	Dim action
	action=Request("action")

	Dim strlsh, strddh, strdwmc, strdmmc, strmh, strsbcj, strjcjxh, strmjxx, strrwlr, ifgbl, ifcbl
	Dim strmtjg, strdxjg, strsxjg, strsjtsl, strqjtsl, strqysd, strmtljcc, strrdogg, ijcfx, iqs
	Dim bpjrb, strjrbxs, strjrbcl, strjrbxx, strmjcl, strbz, dtjhjssj, strgjljcc
	Dim dtrwxdsj, strzz, strjsdb, strbm, strlzr, dtlzrq, bgjfs,bqhgj, strgjxs
	Dim sngmjzf, sngmtbl, ibomzf, itsdzf, itszf, itsxxzlzf, igjzf,sngmtjgbl, sngdxjgbl
	Dim strxcbh, strdxqg, strtslb, strcnts, bbtiao,strckdm,strfzxs, strmtrw, strdxrw
	Dim igjfs1, igjfs2, igjfs3, igjfs4, strjgzz, strsjzz, dtjhkssj, dtjgjssj, strdedm, strdefz, strdemt, strdedx

	'���������б�����ʼ��
	strlsh=Trim(Request("lsh")) : strddh=Trim(Request("ddh")) : strdwmc=Trim(Request("dwmc")) : strdmmc=Trim(Request("dmmc"))
	strmh=Trim(Request("mh")) : strsbcj=Trim(Request("sbcj")) : strjcjxh=Trim(Request("jcjxh")) : strmjxx=Request("mjxx")
	strrwlr=Request("rwlr") : strmtjg=Trim(Request("mtjg")) : strdxjg=Trim(Request("dxjg")) : strsxjg=Trim(Request("sxjg"))
	strsjtsl=Trim(Request("sjtsl")) : strqjtsl=Trim(Request("qjtsl")) : strqysd=Trim(Request("qysd")) : strmtljcc=Trim(Request("mtljcc"))
	strgjljcc=Trim(Request("gjljcc")) : strrdogg=Trim(Request("rdogg")) : ijcfx=Trim(Request("jcfx")) : iqs=Request("qs") : bpjrb=Request("pjrb")
	strjrbxs=Request("jrbxs") : strjrbcl=Request("jrbcl") : strjrbxx=Trim(Request("jrbxx")) : strmjcl=Trim(Request("mjcl"))
	strbz=Request("bz") : strckdm=Request("ckdm") : strfzxs=Request("fzxs") : dtjhkssj=Request("jhkssj") : dtjhjssj=Request("jhjssj")
	dtjgjssj=Request("jgjssj") :  dtrwxdsj=now() : dtlzrq=now() : sngmtjgbl=Request("mtjgbl") : sngdxjgbl=Request("dxjgbl")
	strzz=Request("zz") : strjsdb=Request("jsdb") : strbm=session("user_depart") : strlzr=session("userName")
	sngmjzf=Request("mjzf") : sngmtbl=Request("mtbl") : ibomzf=Request("bomzf") : itsdzf=Request("tsdzf")
	itszf=Request("tszf") : itsxxzlzf=Request("tsxxzlzf") : igjzf=Request("gjzf")
	strtslb=Request("tslb") : strxcbh=Request("xcbh") : strcnts=Request("cnts") : bbtiao=Request("beit")
	strjgzz=Request("jgzz") : strsjzz=Request("sjzz") : strdxqg=Request("dxqg")
	igjfs1=Request("ssgjf") : igjfs2=Request("qbfgjf")
	igjfs3=Request("qgjf") : igjfs4=Request("hgjf")
	strmtrw=Request("mtrw") : strdxrw=Request("dxrw") : strdedm=Request("dedm") : strdefz=Request("defz")		

	sngmjzf=NullToNum(sngmjzf)
	sngmtbl=NulltoNum(sngmtbl)
	sngmtjgbl=NulltoNum(sngmtjgbl)
	sngdxjgbl=NulltoNum(sngdxjgbl)
	ibomzf=NulltoNum(ibomzf)
	itsdzf=NulltoNum(itsdzf)
	itszf=NulltoNum(itszf)
	itsxxzlzf=NulltoNum(itsxxzlzf)
	igjzf=NulltoNum(igjzf)
	strdemt=NulltoNum(strdemt)
	strdedx=NulltoNum(strdedx)

	'��ʼ��ģ�߶�����Ϣ
	strSql="select * from [c_fzbl]"
	set rs=xjweb.Exec(strSql, 1)
	ifgbl=CSng(rs("fgbl"))
	ifcbl=CSng(rs("fcbl"))
	rs.close

	Select Case strmjxx
		Case "ģͷ"
			strdefz=strdefz*0.4
			sngmtbl=100
		Case "����"
			strdefz=strdefz*0.6
			sngmtbl=0
	End select	
	
	If igjfs2<>0 Then
		strdemt=strdefz*strfzxs*sngmtbl/100 + 400
	else
		strdemt=strdefz*strfzxs*sngmtbl/100
	End If
	strdedx=strdefz*strfzxs*(100-sngmtbl)/100
	If igjfs3<>0 Then
		strdemt=strdemt + 200*sngmtbl/100
		strdedx=strdedx + 200*(100-sngmtbl)/100
	End If	
	Select Case strmtrw
		Case ""
			strdemt=0
		Case "���"
			strdemt=Round(strdemt,1)
		Case "����"
			strdemt=Round(strdemt*ifgbl,1)
		Case "����"
			strdemt=Round(strdemt*ifcbl,1)
	End select
	Select Case strdxrw
		Case ""
			strdedx=0
		Case "���"
			strdedx=Round(strdedx,1)
		Case "����"
			strdedx=Round(strdedx*ifgbl,1)
		Case "����"
			strdedx=Round(strdedx*ifcbl,1)
	End select	

	'�Դ�������ݽ��д���
	strMsg=""
	If strlsh="" Then strMsg="��ˮ��Ϊ��!<br>"
	If strddh="" Then strMsg=strMsg & "������Ϊ��!<br>"
	If strdwmc="" Then strMsg=strMsg & "�ͻ�����Ϊ��!<br>"
	If strdmmc="" Then strMsg=strMsg & "��������Ϊ��!<br>"
	If strmh=""  Then strMsg=strMsg & "ģ��Ϊ��!<br>"
	If strsbcj=""  Then strMsg=strMsg & "�豸����Ϊ��!<br>"
	If strjcjxh=""  Then strMsg=strMsg & "�������ͺ�Ϊ��!<br>"
	If strqysd=""  Then strMsg=strMsg & "ǣ���ٶ�Ϊ��!<br>"
	If strckdm=""  Then strMsg=strMsg & "�ο����治��Ϊ��!<br>"
	If bbtiao=""  Then strMsg=strMsg & "���ڲ�����ʱ�Ƿ񱱵�����Ϊ��!<br>"
	If strcnts="true" and strtslb=""  Then strMsg=strMsg & "���ڵ���ʱ���������Ϊ��!<br>"
	If strxcbh=""  Then strMsg=strMsg & "�Ͳıں�Ϊ��!<br>"
	If strmjxx<>"ģͷ" And strdxjg="" Then strMsg=strMsg & "���ͽṹΪ��!<br>"
	If strmjxx<>"ģͷ" And strsxjg="" Then strMsg=strMsg & "ˮ��ṹΪ��!<br>"
	If strmtljcc=""  Then strMsg=strMsg & "ģͷ���ӳߴ�Ϊ��!<br>"
	If strrdogg=""  Then strMsg=strMsg & "�ȵ�ż���Ϊ��!<br>"
	If strmjcl=""  Then strMsg=strMsg & "ģ�߲���Ϊ��!<br>"
	If sngmjzf=0 Then strMsg=strMsg & "ģ���ܷ�Ϊ��!<br>"
	'If sngmtbl=0 Then strMsg=strMsg & "ģͷ����Ϊ��!<br>"
	If sngmtjgbl=0 Then strMsg=strMsg & "ģͷ�ṹ����Ϊ��!<br>"
	If sngdxjgbl=0 Then strMsg=strMsg & "���ͽṹ����Ϊ��!<br>"
	If strdemt=0 and strdedx=0 Then strMsg=strMsg & "ģͷ�Ͷ��Ͷ����ͬʱΪ0!<br>"
	If strzz=""  Then
		If strjgzz="" or strsjzz="" Then strMsg=strMsg & "�鳤û��ѡ��!<br>"
	End If
	If strjsdb=""  Then strMsg=strMsg & "��������û��ѡ��!<br>"

	If strMsg <> "" and action <> "BzChan" Then
		infoTitle="���ݲ�����"
		infoContents=strMsg & "<br>���<a href=""#"" onclick='history.go(-1);'>����ǰҳ</a>��������"
		GotoPrompt()
	End If

	If Not ChkStr(strlsh) and action <> "BzChan" Then Call JsAlert("��ˮ���к��зǷ��ַ���\n��ո񡢻س��������š�˫���ŵȣ�","")

	'���鳤�ó��ǵڼ���
'	Dim igroup
'	igroup=0
'	strSql="select [user_group] from [ims_user] where [user_name]='"&strzz&"'"
'	Set Rs=xjweb.exec(strSql, 1)
'	igroup = Rs("user_group")

	'������⺯�������￪ʼ
	Select Case action
		Case "add"
			Call mtask_add()
		Case "change"
			Call mtask_change()
		Case "BzChan"
			Call Bz_Chan()
		Case else
			response.write "action=" & action
	End select

	'������������
	Function mtask_add()
		'�����ˮ���Ƿ��Ѵ���
		Set Rs=xjweb.exec("select lsh from [mtask] where [lsh]='"&strlsh&"'",1)
		If Not(Rs.eof Or Rs.bof) Then
			Call JsAlert("��ˮ�� " & strlsh & " �������Ѵ���!�������ˮ��!","")
			Exit Function
		End If
		Rs.Close
		'Response.End
		strSql="select * from [mtask]"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Rs.AddNew
			Rs("ddh")=strddh
			Rs("lsh")=strlsh
			Rs("dwmc")=strdwmc
			Rs("dmmc")=strdmmc
			Rs("mh")=strmh
			Rs("sbcj")=strsbcj
			Rs("jcjxh")=strjcjxh
			If strmtjg<>"" Then Rs("mtjg")=strmtjg
			If strdxjg<>"" Then Rs("dxjg")=strdxjg
			If strsxjg<>"" Then Rs("sxjg")=strsxjg
			Rs("sjtsl")=strsjtsl
			Rs("qjtsl")=strqjtsl
			Rs("qysd")=strqysd
			Rs("mtljcc")=strmtljcc
			Rs("gjljcc")=strgjljcc
			Rs("rdogg")=strrdogg
			Rs("mjxx")=strmjxx
			Rs("rwlr")=strrwlr			
			Rs("mtrw")=strmtrw
			Rs("dxrw")=strdxrw
			Rs("ckdm")=strckdm
			Rs("dedm")=strdedm
			Rs("demt")=strdemt
			Rs("dedx")=strdedx
			Rs("fzxs")=strfzxs
			Rs("cnts")=strcnts
			Rs("beit")=bbtiao
			If strtslb<>"" Then Rs("tslb")=strtslb
			Rs("xcbh")=strxcbh
			Rs("dxqg")=strdxqg
			Rs("jcfx")=ijcfx
			Rs("qs")=iqs
			Rs("pjrb")=bpjrb
			Rs("jrbxs")=strjrbxs
			Rs("jrbcl")=strjrbcl
			If strjrbxx<>"" Then Rs("jrbxx")=strjrbxx
			Rs("mjcl")=strmjcl
			If strbz<>"" Then Rs("bz")=strbz
			Rs("rwxdsj")=dtrwxdsj
			Rs("jhkssj")=dtjhkssj
			Rs("jhjgsj")=dtjgjssj
			Rs("jhjssj")=dtjhjssj
'			Rs("zz")=strzz
			Rs("jgzz")=strjgzz
			Rs("sjzz")=strsjzz
'			Rs("group")=igroup
			Rs("jsdb")=strjsdb
			Rs("bm")=strbm
			Rs("lzr")=strlzr
			Rs("lzrq")=dtlzrq
'			Rs("gjfs")=bgjfs
			Rs("SSGJ")=igjfs1
			Rs("QBFGJ")=igjfs2
			Rs("QGJ")=igjfs3
			Rs("HGJ")=igjfs4
'			Rs("qhgj")=bqhgj
			Rs("mjzf")=sngmjzf
			Rs("mtbl")=sngmtbl
			Rs("mtjgbl")=sngmtjgbl
			Rs("dxjgbl")=sngdxjgbl
			Rs("bomzf")=ibomzf
			Rs("tsdzf")=itsdzf
			Rs("tszf")=itszf
			Rs("tsxxzlzf")=itsxxzlzf
			Rs("gjzf")=igjzf
		Rs.update
		Rs.Close
		Call JsAlert("��������ӳɹ�!", "mtask_add.asp")
		Response.End
	End Function

	'�������������
	Function mtask_change()
		'�����ˮ���Ƿ��Ѵ���
		Dim iid
		iid=Request("id")
		Set Rs=xjweb.Exec("select lsh from [mtask] where [lsh]='"&strlsh&"' And id<>"&iid&" ",1)
		If Not(Rs.eof Or Rs.bof) Then
			Call JsAlert("��ˮ�� " & strlsh & " �������Ѵ���! �������ˮ��!","")
			Exit Function
		End If
		Rs.Close

		strSql="select * from [mtask] where [id]=" & iid
		Call xjweb.exec("",-1)
		strMsg="����������"
		Rs.open strSql,Conn,1,3
			Rs("ddh")=strddh
			Rs("lsh")=strlsh
			Rs("dwmc")=strdwmc
			Rs("dmmc")=strdmmc
			Rs("mh")=strmh
			Rs("sbcj")=strsbcj
			Rs("jcjxh")=strjcjxh
			If strmtjg<>"" Then Rs("mtjg")=strmtjg
			If strdxjg<>"" Then Rs("dxjg")=strdxjg
			If strsxjg<>"" Then Rs("sxjg")=strsxjg
			Rs("sjtsl")=strsjtsl
			Rs("qjtsl")=strqjtsl
			Rs("qysd")=strqysd
			Rs("mtljcc")=strmtljcc
			Rs("gjljcc")=strgjljcc
			Rs("rdogg")=strrdogg
			Rs("mjxx")=strmjxx
			Rs("rwlr")=strrwlr			
			Rs("mtrw")=strmtrw
			Rs("dxrw")=strdxrw
			Rs("ckdm")=strckdm
			Rs("dedm")=strdedm
			Rs("demt")=strdemt
			Rs("dedx")=strdedx
			Rs("fzxs")=strfzxs
			Rs("cnts")=strcnts
			Rs("beit")=bbtiao
			If strtslb<>"" Then Rs("tslb")=strtslb
			Rs("xcbh")=strxcbh
			Rs("dxqg")=strdxqg
			Rs("jcfx")=ijcfx
			Rs("qs")=iqs
			Rs("pjrb")=bpjrb
			Rs("jrbxs")=strjrbxs
			Rs("jrbcl")=strjrbcl
			If strjrbxx<>"" Then Rs("jrbxx")=strjrbxx
			Rs("mjcl")=strmjcl
			If strbz<>"" Then Rs("bz")=strbz
			Rs("rwxdsj")=dtrwxdsj
			Rs("jhkssj")=dtjhkssj
			Rs("jhjgsj")=dtjgjssj
			Rs("jhjssj")=dtjhjssj
			Rs("zz")=strzz
			Rs("jgzz")=strjgzz
			Rs("sjzz")=strsjzz
'			Rs("group")=igroup
			Rs("jsdb")=strjsdb
			Rs("bm")=strbm
			Rs("lzr")=strlzr
			Rs("lzrq")=dtlzrq
			Rs("gjfs")=0
			Rs("qhgj")=0
			Rs("SSGJ")=igjfs1
			Rs("QBFGJ")=igjfs2
			Rs("QGJ")=igjfs3
			Rs("HGJ")=igjfs4
			Rs("mjzf")=sngmjzf
			Rs("mtbl")=sngmtbl
			Rs("mtjgbl")=sngmtjgbl
			Rs("dxjgbl")=sngdxjgbl
			Rs("bomzf")=ibomzf
			Rs("tsdzf")=itsdzf
			Rs("tszf")=itszf
			Rs("tsxxzlzf")=itsxxzlzf
			Rs("gjzf")=igjzf
		Rs.update
		Rs.Close
		Call JsAlert("��ˮ�� " & strlsh &  " ��������ĳɹ�!", "mtask_change.asp")
		Response.End
	End Function

	'���ı�ע
	Function Bz_Chan()
		'�����ˮ���Ƿ��Ѵ���
		Dim iid
		iid=Request("id")
		strSql="select * from [mtask] where [id]=" & iid
		Call xjweb.exec("",-1)
		strMsg="����������"
		Rs.open strSql,Conn,1,3
			If strbz<>"" Then Rs("bz")=strbz
		Rs.update
		Rs.Close
		Call JsAlert("��ע���ĳɹ�!", "mtask_change.asp")
		Response.End
	End Function
%>
