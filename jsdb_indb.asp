<!--#include file="include/conn.asp"-->
<%
'2013/1/14 10:15
	'��Ӽ����������������	
	Dim action, strhth, strkhmc, strdrwnr, strfz, strjhsj, strzz
	action=Request("action")
	strhth=Trim(Request("hth")) : strkhmc=Trim(Request("khmc")) : strdrwnr=Trim(Request("rwnr")) : strjhsj=Trim(Request("jhjssj"))
	strzz=Trim(Request("sjr")) : strfz=NulltoNum(Request("jcf"))

	Select Case action
		Case "add"
			Call jsdb_add()
		Case "change"
			Call jsdb_change()
	End select

	'������������
	Function jsdb_add()
	'�Դ�������ݽ��д���
	Dim strMsg
	strMsg=""
	If strhth="" Then strMsg="��ͬ��Ϊ��!<br>"
	If strkhmc="" Then strMsg=strMsg & "�ͻ�����Ϊ��!<br>"
	If strzz="" Then strMsg=strMsg & "�鳤Ϊ��!<br>"
	If strfz=0 Then strMsg=strMsg & "��ֵΪ0!<br>"

	If strMsg <> ""Then
		infoTitle="���ݲ�����"
		infoContents=strMsg & "<br>���<a href=""#"" onclick='history.go(-1);'>����ǰҳ</a>��������"
		GotoPrompt()
	End If
	
		'����ͬ���Ƿ��Ѵ���
		strSql="select * from [jsdb] where [hth]='" & strhth & "'"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		If Not(Rs.eof Or Rs.bof) Then
			If IsNull(Rs("shjssj")) Then
				Rs("khmc")=strkhmc
				If strdrwnr<>"" Then Rs("rwnr")=strdrwnr
				Rs("zz")=strzz
				Rs("jcf")=strfz
				Rs("jhjssj")=strjhsj
				Rs.update
				Rs.Close
				Call JsAlert("������ĳɹ�!", "jsdb_add.asp")
			else
				Rs.Close
				Call JsAlert("��������ɣ��޷��޸�!!", "jsdb_add.asp")
			End If
		End If
		Rs.Close
		
		strSql="select * from [jsdb]"
		Call xjweb.exec("",-1)
		Rs.open strSql,Conn,1,3
		Rs.AddNew
			Rs("hth")=strhth
			Rs("khmc")=strkhmc
			If strdrwnr<>"" Then Rs("rwnr")=strdrwnr
			Rs("zz")=strzz
			Rs("jcf")=strfz
			Rs("jhjssj")=strjhsj
		Rs.update
		Rs.Close
		Call JsAlert("��������ӳɹ�!", "jsdb_add.asp")
	End Function
%>
