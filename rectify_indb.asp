<!--#include file="include/conn.asp"-->
<%
	Call ChkAble(11)
	Dim action, strzrbm, strbh, strxxbm, strjssj, strbhgnr, stryfcsyq, strqxsj, strps,  stryyfx, strjzcs, strlsqk, stryzjl, strwczk, iid
	action="" : strzrbm="" : strbh="" : strxxbm="" : strjssj="" : strbhgnr="" : stryfcsyq="" : strqxsj="" : strps="" : stryyfx="" : strjzcs="" : strlsqk="" : stryzjl="" : strwczk="������" : iid=0
	action=LCase(Request("action"))
	strzrbm=Request("zrbm")
	strbh=Request("bh")
	strxxbm=Request("xxbm")
	strjssj=Request("jssj")
	strbhgnr=Request("bhgnr")
	stryfcsyq=Request("yfcsyq")
	strqxsj=Request("qxsj")
	strps=Request("ps")
	stryyfx=Request("yyfx")
	strjzcs=Request("jzcs")
	strlsqk=Request("lsqk")
	stryzjl=Request("yzjl")
	If IsNumeric(Trim(Request("id"))) Then iid=CLng(Trim(Request("id")))

	'������⺯�������￪ʼ
	Select Case action
		Case "add"
			If strzrbm="" Or strjssj="" Or strbhgnr="" Then
				Call JsAlert("��ȷ����Ϣ��������!�����ȷ����ڽ���!","Rectify_list.asp")
			Else
				Call Rectify_Add()
			End If
		Case "change"
			If strzrbm="" Or strbhgnr="" Or iid="" Then
				Response.Write strbhgnr
				Call JsAlert("��ȷ����Ϣ��������!","")
			Else
				Call Rectify_Change()
			End If
		Case "delete"
			If iid=0 Then
				Call JsAlert("��ȷ�ϴ�ϵͳ��ڽ���!","")
			Else
				strSql="delete from [rectify] where id=" & iid
				Call xjweb.Exec(strSql, 0)
				Call JsAlert("�ⲿ������Ϣɾ���ɹ�","Rectify_list.asp")
			end if
		Case Else
			Call JsAlert("action="&action&", ����ϵ����Ա!","Rectify_list.asp")
	End Select
	
	'��֤�ⲿ��Ϣ������
	Function Rectify_zt()
		If stryfcsyq <> "" Then strwczk = "������"		 
		If stryyfx <> "" Then strwczk = "�ƶ���ʩ��" 
		If strjzcs <> "" Then strwczk = "���ټ����" 
		If strlsqk <> "" Then strwczk = "��֤��"  
		If stryzjl <> "" Then strwczk = "�ѱջ�"  
	End Function
	
	'�ⲿ������Ϣ ���
	Function Rectify_Add()
		strSql="select * from [Rectify]"
		Call xjweb.Exec("",-1)
		Call Rectify_zt()
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("zrbm")=strzrbm
			Rs("bh")=strbh
			Rs("xxbm")=strxxbm
			Rs("jssj")=strjssj
			Rs("bhgnr")=strbhgnr
			Rs("yfcsyq")=stryfcsyq
			Rs("qxsj")=strqxsj
			Rs("ps")=strps
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl
			Rs("wczk")=strwczk
		Rs.Update
		Rs.Close
		Call JsAlert("����/Ԥ����ʩ����ӳɹ�","Rectify_add.asp")
	End Function

	'�����ⲿ������Ϣ ���	
	Function Rectify_Change()
		'���ID���Ƿ����
		Set Rs=xjweb.Exec("select * from [Rectify] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("��� " & iid & " ����/Ԥ����ʩ�����ڣ�","Rectify_list.asp")
			Rs.Close
			Exit Function
		End If
		Rs.Close

		strSql="select * from [Rectify] where id=" & iid
		Call xjweb.Exec("",-1)
		Call Rectify_zt()
		'strmsg="���ݿ����"
		Rs.open strSql,conn,1,3
			Rs("zrbm")=strzrbm
			Rs("bh")=strbh
			Rs("xxbm")=strxxbm
'			Rs("jssj")=strjssj
			Rs("bhgnr")=strbhgnr
			Rs("yfcsyq")=stryfcsyq
			Rs("qxsj")=strqxsj
			Rs("ps")=strps
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl
			Rs("wczk")=strwczk
		Rs.update
		Rs.close

		Call JsAlert("����/Ԥ����ʩ����ĳɹ�","Rectify_list.asp")
	End Function
%>