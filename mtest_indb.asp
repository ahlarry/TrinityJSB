<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble("1,6")
	dim action, strtsyy, strtslr, btsps, strlsh, iid
	action="" : strtsyy="" : strtslr="" : btsps=false : strlsh=""
	action=request("action")
	strtsyy=request("tsyy")
	strtslr=trim(request("tslr"))
	btsps=request("tsps")
	strlsh=request("lsh")
	iid=request("id")

	'������⺯�������￪ʼ
	select case action
		case "add"
			if strtsyy="" or strtslr="" or strlsh="" then
				Call JsAlert("��ȷ�ϴ���ȷ��ڽ��벢��֤��Ϣ��������!","mtest_add.asp")
			else
				if btsps then
					Call mtestps_add()
				else
					Call mtest_add()
				end if
			end if
		case "change"
			if strtsyy="" or strtslr="" or not isnumeric(iid) then
				Call JsAlert("��ȷ�ϴ���ȷ��ڽ��벢��֤��Ϣ��������!","mtest_add.asp")
			else
				Call mtest_change()
			end if
		case "delete"
			if not(isnumeric(iid)) then
				Call JsAlert("��ȷ�ϴ���ȷ��ڽ��벢��֤��Ϣ��������!","mtest_add.asp")
			else
				Dim bPs
				strSql="delete from [ts_tsxx] where id=" & iid
				Set Rs=xjweb.Exec("Select lsh,ps from [ts_tsxx] where id=" & iid,1)
				'��¼������Ϣ�������Ϣ
				strlsh=Rs("lsh")
				bPs=Rs("ps")
				Rs.Close
				Call xjweb.Exec(strSql, 0)
				'���ĵ��ԵĴ���
				strSql="update [ts_mould] set tscs="&xjweb.RsCount("[ts_tsxx] where lsh='"&strlsh&"' and not ps")&" where lsh='"&strlsh&"'"
				Call xjweb.Exec(strSql,0)
				If bPs Then
					Call JsAlert("����������Ϣɾ���ɹ�!","mtest_display.asp?s_lsh="&strlsh&"")
				Else
					Call JsAlert("ģ�ߵ�����Ϣɾ���ɹ�!","mtest_display.asp?s_lsh="&strlsh&"")
				End If
				response.end
			end if
		case else
			response.write "action=" & action
	end select

	'������Ϣ���
	function mtest_add()
		strSql="select * from ts_tsxx"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		rs.addnew
			rs("lsh")=strlsh
			rs("tsyy")=strtsyy
			rs("tslr")=strtslr
			rs("tssj")=now()
			rs("tsr")=Session("userName")
		rs.update
		rs.close

		'����Ϣ��ts_mould ��
		strSql="select * from [ts_mould] where lsh='"&strlsh&"'"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
			if isnull(rs("tskssj")) then rs("tskssj")=now()
			rs("tscs")=rs("tscs") + 1
			rs("tsgxsj")=now()
		rs.update
		rs.close
		Call JsAlert("��ˮ�� ��" & strlsh & "�� ģ�ߵ�����Ϣ��ӳɹ�!","mtest_display.asp?s_lsh="&strlsh&"")
	end function

	function mtestps_add()
		strSql="select * from ts_tsxx"
		Call xjweb.Exec("",-1)
		rs.open strSql,conn,1,3
		rs.addnew
			rs("lsh")=strlsh
			rs("tsyy")=strtsyy
			rs("tslr")=strtslr
			rs("tssj")=now()
			rs("tsr")=Session("userName")
			rs("ps")=true
		rs.update
		rs.close
		Call JsAlert("��ˮ�� ��" & strlsh & "�� ģ�ߵ���������ӳɹ�!","mtest_display.asp?s_lsh="&strlsh&"")
	end function

	'���������������
	function mtest_change()
		'�����ˮ���Ƿ��Ѵ���
		set rs=xjweb.Exec("select * from [ts_tsxx] where id="&iid,1)
		if rs.eof or rs.bof then
			Call JsAlert("ID��Ϊ ��"&iid&"�� �ĵ�����Ϣ������!","mtest_list.asp")
			exit function
		End if
		rs.close

		strSql="select * from ts_tsxx where id=" & iid
		Call xjweb.Exec("",-1)
		strmsg="���ݿ����"
		rs.open strSql,conn,1,3
			rs("tsyy")=strtsyy
			rs("tslr")=strtslr
		rs.update
		rs.close

		'strSql="insert into ims_log (loguser, logip, logtopic, loginfo, logtime) values ('"&session("userName")&"','"&request.servervariables("local_addr")&"','����������','"&strmsg&"','"&now()&"')"
		'Call xjweb.Exec(strSql,0)
		Call JsAlert("������Ϣ���ĳɹ�!","mtest_display.asp?s_lsh="&strlsh&"")
	end function
%>