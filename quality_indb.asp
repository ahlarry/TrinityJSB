<!--#include file="include/conn.asp"-->
<%
	Call ChkAble(11)
	Dim action, strkhmc, strlxr, strlxdh, strhth, strgzlh, strjssj, strZrr, strzywt, stryjcs, stryyfx, strjzcs, strlsqk, stryzjl, strwczk, iid
	action="" : strkhmc="" : strlxr="" : strlxdh="" : strhth="" : strgzlh="" : strjssj="" : strZrr="" : strzywt="" : stryjcs="" : stryyfx="" : strjzcs="" : strlsqk="" : stryzjl="": strwczk="������" : iid=0
	action=LCase(Request("action"))
	strkhmc=Request("khmc")
	strlxr=Request("lxr")
	strlxdh=Request("lxdh")
	strhth=Request("hth")
	strgzlh=Request("gzlh")
	strjssj=Request("jssj")
	strZrr=Request("zrr")
	strzywt=Request("zywt")
	stryjcs=Request("yjcs")
	stryyfx=Request("yyfx")
	strjzcs=Request("jzcs")
	strlsqk=Request("lsqk")
	stryzjl=Request("yzjl")
	If IsNumeric(Trim(Request("id"))) Then iid=CLng(Trim(Request("id")))

	'������⺯�������￪ʼ
	Select Case action
		Case "add"
			If strhth="" Or strkhmc="" Or strZrr="" Or strzywt="" Then
				Call JsAlert("��ȷ����Ϣ��������!�����ȷ����ڽ���!","quality_list.asp")
			Else
				Call quality_Add()
			End If
		Case "change"
			If strhth="" Or strkhmc="" Or strZrr="" Or strzywt="" Or iid="" Then
				Call JsAlert("��ȷ����Ϣ��������!","")
			Else
				Call quality_Change()
			End If
		Case "delete"
			If iid=0 Then
				Call JsAlert("��ȷ�ϴ�ϵͳ��ڽ���!","")
			Else
				strSql="delete from [quality] where id=" & iid
				Call xjweb.Exec(strSql, 0)
				Call JsAlert("�ⲿ������Ϣɾ���ɹ�","quality_list.asp")
			end if
		Case Else
			Call JsAlert("action="&action&", ����ϵ����Ա!","quality_list.asp")
	End Select
	
	'��֤�ⲿ��Ϣ������
	Function quality_zt()
		If stryjcs <> "" Then strwczk = "������"		 
		If stryyfx <> "" Then strwczk = "�ƶ���ʩ��" 
		If strjzcs <> "" Then strwczk = "���ټ����" 
		If strlsqk <> "" Then strwczk = "��֤��"  
		If stryzjl <> "" Then strwczk = "�ѱջ�"  
	End Function
	
	'�ⲿ������Ϣ ���
	Function quality_Add()
		strSql="select * from [quality]"
		Call xjweb.Exec("",-1)
		Call quality_zt()
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("khmc")=strkhmc
			Rs("lxr")=strlxr
			Rs("lxdh")=strlxdh
			Rs("hth")=strhth
			If strgzlh<>"" Then Rs("gzlh")=strgzlh
			Rs("jssj")=strjssj
			Rs("zrr")=strZrr
			Rs("zywt")=strzywt
			Rs("yjcs")=stryjcs
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl
			Rs("wczk")=strwczk
		Rs.Update
		Rs.Close
		Call JsAlert("�ⲿ������Ϣ��ӳɹ�","quality_add.asp")
	End Function

	'�����ⲿ������Ϣ ���	
	Function quality_Change()
		'���ID���Ƿ����
		Set Rs=xjweb.Exec("select * from [quality] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("ID�� " & iid & " ������������ڣ�","quality_list.asp")
			Rs.Close
			Exit Function
		End If
		Rs.Close

		strSql="select * from [quality] where id=" & iid
		Call xjweb.Exec("",-1)
		Call quality_zt()
		'strmsg="���ݿ����"
		Rs.open strSql,conn,1,3
			Rs("khmc")=strkhmc
			Rs("lxr")=strlxr
			Rs("lxdh")=strlxdh
			Rs("hth")=strhth
			Rs("gzlh")=strgzlh
			Rs("jssj")=strjssj
			Rs("zrr")=strZrr
			Rs("zywt")=strzywt
			Rs("yjcs")=stryjcs
			Rs("yyfx")=stryyfx
			Rs("jzcs")=strjzcs
			Rs("lsqk")=strlsqk
			Rs("yzjl")=stryzjl	
			Rs("wczk")=strwczk			
		Rs.update
		Rs.close

		Call JsAlert("�ⲿ������Ϣ���ĳɹ�","quality_list.asp")
	End Function
%>