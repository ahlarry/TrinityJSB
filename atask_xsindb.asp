<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble(4)
	If Session("userGroup") <> 5 Then Call JsAlert("����ϵ�����鳤���д˸����������!","atask.asp")
	dim str_yxs, str_newxs, str_lsh, TmpSql,TmpRs
	str_yxs="" : str_newxs="" : str_lsh=""

	str_yxs=Trim(Request("yxs"))
	str_newxs=Trim(Request("newxs"))
	str_lsh=Trim(Request("ylsh"))

	if str_lsh="" then
		Call JsAlert("��ˮ�Ų�ȷ��,�޷�ȷ������ĵ�������!\n���������ڽ���!","atask_changexs.asp")
	end if

	set rs=xjweb.Exec("select lsh from [mantime] where lsh='"&str_lsh&"'",1)
	if rs.eof or rs.bof then
		Call JaAlert("��ˮ�� ��"&strlsh&"�� ���񲻴���!���ʵ!","atask_changexs.asp")
	end if
	Rs.close
	
	TmpSql="select * from [mantime] where (rwlr='ȫ�׵��Ժϸ�' or rwlr='ģͷ���Ժϸ�' or rwlr='���͵��Ժϸ�' or rwlr like '%����%' or rwlr like '%����%' or rwlr like '%����%') and lsh='"&str_lsh&"' and jc="&str_yxs&""
	Set TmpRs=Server.CreateObject("adodb.recordset")
	TmpRs.open TmpSql,conn,1,3
	Do while not TmpRs.Eof
		If str_newxs<>"" and str_newxs<>str_yxs Then 
			TmpRs("jc")=str_newxs
			If IsNull(TmpRs("rwfz")) Then TmpRs("rwfz")=Round(TmpRs("fz")/str_yxs,1)
			TmpRs("fz")=Round(TmpRs("rwfz")*str_newxs,1)
			TmpRs("xgsj")=now()
			TmpRs("bz")=session("userName")&"("&request.servervariables("local_addr")&")||�޸�ϵ��:"&str_yxs&"��"&str_newxs
		End If
	TmpRs.update
	TmpRs.movenext
	Loop
	TmpRs.close
Call JsAlert(str_lsh&"���Է�ֵϵ�����ĳɹ���","atask_changexs.asp")
%>
