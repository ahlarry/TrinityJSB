<!--#include file="include/conn.asp"-->
<%
	dim action, strrwlx, strrwlr, strzrr, izf, iid,strxldh,strxlxh,stryhdw,strgzyy,   strzbfa,strmjmc,strjssj,strzrfp,strylsh
	action="" : strrwlx="" : strrwlr="" : strzrr="" : izf=0 :  iid=0 :  strxldh="" : strxlxh="" : stryhdw="" : strgzyy="" :  strzbfa="" : strmjmc="":strjssj="":strzrfp="" : strylsh=""
	action=request("action")
	strrwlx=request("rwlx")
	strxldh=request("xldh")
	strxlxh=request("xlxh")
	stryhdw=request("yhdw")
	strzrfp=Request("zrfp")
	strylsh=trim(request("ylsh"))
	strgzyy=trim(request("gzyy"))
	strzbfa=trim(request("zbfa"))
	strmjmc=trim(request("mjmc"))
	iid=clng(request("id"))
	Call SelectT()

	'������⺯�������￪ʼ
	select case action

		case "add"
			if strrwlx="" or strrwlr="" or strzrr="" or izf=0 then
				Call JsAlert("��ȷ����Ϣ��������!","")
			else
				call ftask_add()
			end if
		case "change"
			if strrwlx="" or strrwlr="" or strzrr="" or izf=0 or not(isnumeric(iid)) then
				Call JsAlert("��ȷ����Ϣ��������!","")
			else
				call ftask_change()
			end if
		case "delete"
			if not(isnumeric(iid)) then
				Call JsAlert("��ȷ�ϴ�ϵͳ��ڽ���!","")
			else
				strSql="delete from [ftask] where id=" & iid
				call xjweb.Exec(strSql, 0)
				Call JsAlert("��������ɾ���ɹ�","ftask_list.asp")
			end if
		case else
			Call JsAlert("action="&action&", ����ϵ����Ա!","ftask_list.asp")
	end select

	'�������������
	Function ftask_Add()
		Dim tmpSql, tmpRs, iGroup
		tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strzrr&"'"
		Set tmpRs=xjweb.Exec(tmpSql,1)
		If Not(tmpRs.Eof Or tmpRs.Bof) Then
			iGroup=tmpRs("user_group")
		Else
			iGroup=0
		End If
		tmpRs.Close
	
		strSql="select * from [ftask]"
		call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		rs.addnew
			rs("rwlx")=strrwlx
			if strxldh<>"" then rs("xldh")=strxldh end if
			rs("rwlr")=strrwlr
			rs("zrr")=strzrr
			rs("xz")=iGroup
			rs("zf")=izf
			rs("jssj")=strjssj
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update
		rs.close
		Call JsAlert("����������ӳɹ�","ftask_add.asp")
	end function

	'���������������
	function ftask_change()
		'�����ˮ���Ƿ��Ѵ���
		set rs=xjweb.Exec("select * from [ftask] where id="&iid,1)
		if rs.eof or rs.bof then
			Call JsAlert("ID�� " & iid & " �������񲻴��ڣ�","ftask_list.asp")
			Exit Function
		end if
		rs.close
		Dim tmpSql, tmpRs, iGroup
		tmpSql="Select [user_group] from [ims_user] where [user_name]='"&strzrr&"'"
		Set tmpRs=xjweb.Exec(tmpSql,1)
		If Not(tmpRs.Eof Or tmpRs.Bof) Then
			iGroup=tmpRs("user_group")
		Else
			iGroup=0
		End If
		tmpRs.Close

		strSql="select * from [ftask] where id=" & iid
		call xjweb.Exec("",-1)
		'strmsg="���ݿ����"
		rs.open strSql,conn,1,3
			rs("rwlx")=strrwlx
			if strxldh<>"" then rs("xldh")=strxldh end if
			rs("rwlr")=strrwlr
			rs("zrr")=strzrr
			rs("xz")=iGroup
			rs("zf")=izf
			rs("jssj")=strjssj
			rs("lzr")=session("userName")
			rs("lzrq")=now()
		rs.update
		rs.close

		'sql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','����������','"&strmsg&"','"&now()&"')"
		'call xjweb.Exec(sql,0)
		Call JsPrompt("����������ĳɹ�")
	End Function
	function SelectT()
	'��������ͬ��ʼ������
	if strrwlx="��������" then
	strrwlr="�û���λ:"&stryhdw&"||ģ������:"&strmjmc&"||����С��:"&strxlxh&"||�������������ԭ��:"&strgzyy&"||׼����ȡ����:"&strzbfa&"||���η���:"&strzrfp&"||ԭ��ˮ��:"&strylsh
	izf=Request("zf1")
	strzrr=Request("zrr1")
	strjssj=Request("psy1") & "��" & request("psm1") & "��" & request("psd1") & "��"
	else
	strrwlr=trim(request("rwlr"))
	izf=request("zf")
	strzrr=request("zrr")
	strjssj=request("psy") & "��" & request("psm") & "��" & request("psd") & "��"
	end if
	end function

%>