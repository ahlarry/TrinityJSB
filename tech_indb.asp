<!--#include file="include/conn.asp"-->
<%
	Call ChkAble(7)
	Dim action, strLsh, strBkmc, strClyj, strZrr, strXxms, strYyfx, strYfcs, iid
	action="" : strLsh="" : strBkmc="" : strClyj="" : strZrr="" : strXxms="" : strYyfx="" :strYfcs="" : iid=0
	action=LCase(Request("action"))
	strLsh=Trim(Request("lsh"))
	strBkmc=Trim(Request("bkmc"))
	strClyj=Request("clyj")
	strZrr=Request("zrr")
	strXxms=Request("xxms")
	strYyfx=Request("yyfx")
	strYfcs=Request("yfcs")
	If IsNumeric(Trim(Request("id"))) Then iid=CLng(Trim(Request("id")))

	'������⺯�������￪ʼ
	Select Case action
		Case "add"
			If strLsh="" Or strBkmc="" Or strClyj="" Or strZrr="" Or strXxms="" Or strYyfx="" Or strYfcs="" Then
				Call JsAlert("��ȷ����Ϣ��������!�����ȷ����ڽ���!","tech_list.asp")
			Else
				Call tech_Add()
			End If
		Case "change"
			If strLsh="" Or strBkmc="" Or strClyj="" Or strZrr="" Or strXxms="" Or strYyfx="" Or strYfcs="" Or iid=0 Then
				Call JsAlert("��ȷ����Ϣ��������!","")
			Else
				Call tech_Change()
			End If
		Case "delete"
			If iid=0 Then
				Call JsAlert("��ȷ�ϴ�ϵͳ��ڽ���!","")
			Else
				strSql="delete from [tecq_question] where id=" & iid
				Call xjweb.Exec(strSql, 0)
				Call JsAlert("�������ɾ���ɹ�","tech_list.asp")
			end if
		Case Else
			Call JsAlert("action="&action&", ����ϵ����Ա!","tech_list.asp")
	End Select

	'�������������
	Function tech_Add()
		strSql="select * from [tecq_question]"
		Call xjweb.Exec("",-1)
		Rs.open strSql,conn,1,3
		Rs.AddNew
			Rs("tecq_lsh")=strLsh
			Rs("tecq_bkmc")=strBkmc
			Rs("tecq_clyj")=strClyj
			Rs("tecq_zrr")=strZrr
			Rs("tecq_xxms")=strXxms
			Rs("tecq_yyfx")=strYyfx
			Rs("tecq_yfcs")=strYfcs
			Rs("tecq_time")=Now()
		Rs.Update
		Rs.Close
		Call JsAlert("���������ӳɹ�","tech_add.asp")
	End Function

	'���������������
	Function tech_Change()
		'���ID���Ƿ����
		Set Rs=xjweb.Exec("select * from [tecq_question] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("ID�� " & iid & " ������������ڣ�","tech_list.asp")
			Rs.Close
			Exit Function
		End If
		Rs.Close

		strSql="select * from [tecq_question] where id=" & iid
		Call xjweb.Exec("",-1)
		'strmsg="���ݿ����"
		Rs.open strSql,conn,1,3
			Rs("tecq_lsh")=strLsh
			Rs("tecq_bkmc")=strBkmc
			Rs("tecq_clyj")=strClyj
			Rs("tecq_zrr")=strZrr
			Rs("tecq_xxms")=strXxms
			Rs("tecq_yyfx")=strYyfx
			Rs("tecq_yfcs")=strYfcs
			Rs("tecq_time")=Now()
		Rs.update
		Rs.close

		'sql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','����������','"&strmsg&"','"&now()&"')"
		'call xjweb.Exec(sql,0)
		Call JsAlert("��������������ĳɹ�","tech_list.asp")
	End Function
%>