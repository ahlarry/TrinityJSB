<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble(0)
	Dim strFeedBack,strZrr, strkpitem, strgzz, strkpjs, strclsh, striPage, iid
	strZrr=Trim(Request("zrr"))
	strkpitem = trim(request("kpitem"))
	strkpjs = trim(request("kpjs"))
	strclsh = trim(request("clsh"))
	strgzz =request("gzz")
	striPage =request("ipage")
	strFeedBack=""
	If strZrr<>"" Then strFeedBack="zrr="&strZrr
	If strkpitem<>"" Then strFeedBack="kpitem="&strkpitem&"&"&strFeedBack
	If strgzz<>0 Then strFeedBack="gzz="&strgzz&"&"&strFeedBack
	If strkpjs<>"" Then strFeedBack="kpjs="&strkpjs&"&"&strFeedBack
	If striPage<>"0" Then strFeedBack="iPage="&striPage&"&"&strFeedBack
	If strFeedBack<>"" Then strFeedBack="?"&strFeedBack
	iid=Request("id")
	If iid="" Or Not isNumeric(iid) Then Call JsAlert("���������ڽ���!","")
	Call kp_delete()
	Function kp_delete()
		dim chginf1, chginf2, strKpIt
		chginf1="" : chginf2="" : strKpIt=""
		'���ID���Ƿ����
		Set Rs=xjweb.Exec("select * from [kp_jsb] where id="&iid,1)
		If Rs.Eof Or Rs.Bof Then
			Call JsAlert("�˼���������Ϣ�����Ѿ�ɾ��!","ygkp_list.asp"&strFeedBack)
			Rs.Close
			Exit Function
		End If

		If Rs("kp_kpr")<>Session("userName") and not(ChkAble(3)) Then
			Call JsAlert("???????","ygkp_list.asp"&strFeedBack)
		End If

		If Not IsNull(Rs("kp_lsh")) Then
			chginf1=Rs("kp_lsh")
			chginf2=Rs("kp_zlid")
			strKpIt=Rs("kp_item")
		End If
		Rs.Close
		If chginf1<>"" Then
			strSql="delete from [kp_jsb] where kp_lsh='"&chginf1&"' and kp_zlid="&chginf2&""
			Call xjweb.Exec(strSql,0)
			If Instr(strKpIt,"�ڶ����")>0 Then
				strSql="delete from [mantime] where lsh='"&chginf1&"' and rwlr like '%���Ժϸ�(%'"
				Call xjweb.Exec(strSql,0)
			End If
		Else
			strSql="delete from [kp_jsb] where id=" & iid
			Call xjweb.Exec(strSql,0)
		End If

		'sql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','����������','"&strmsg&"','"&now()&"')"
		'call xjweb.Exec(sql,0)
		Call JsAlert("Ա������ɾ���ɹ�","ygkp_list.asp"&strFeedBack)
	End Function
%>