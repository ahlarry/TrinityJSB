<!--#include file="include/conn.asp"-->
<%
	Call ChkPageAble(4)
	If Session("userGroup") <> 5 Then Call JsAlert("请联系第五组长进行此辅助任务分配!","atask.asp")
	dim str_yxs, str_newxs, str_lsh, TmpSql,TmpRs
	str_yxs="" : str_newxs="" : str_lsh=""

	str_yxs=Trim(Request("yxs"))
	str_newxs=Trim(Request("newxs"))
	str_lsh=Trim(Request("ylsh"))

	if str_lsh="" then
		Call JsAlert("流水号不确定,无法确定具体的调试任务!\n请从正规入口进入!","atask_changexs.asp")
	end if

	set rs=xjweb.Exec("select lsh from [mantime] where lsh='"&str_lsh&"'",1)
	if rs.eof or rs.bof then
		Call JaAlert("流水号 【"&strlsh&"】 任务不存在!请核实!","atask_changexs.asp")
	end if
	Rs.close
	
	TmpSql="select * from [mantime] where (rwlr='全套调试合格' or rwlr='模头调试合格' or rwlr='定型调试合格' or rwlr like '%精调%' or rwlr like '%验收%' or rwlr like '%初调%') and lsh='"&str_lsh&"' and jc="&str_yxs&""
	Set TmpRs=Server.CreateObject("adodb.recordset")
	TmpRs.open TmpSql,conn,1,3
	Do while not TmpRs.Eof
		If str_newxs<>"" and str_newxs<>str_yxs Then 
			TmpRs("jc")=str_newxs
			If IsNull(TmpRs("rwfz")) Then TmpRs("rwfz")=Round(TmpRs("fz")/str_yxs,1)
			TmpRs("fz")=Round(TmpRs("rwfz")*str_newxs,1)
			TmpRs("xgsj")=now()
			TmpRs("bz")=session("userName")&"("&request.servervariables("local_addr")&")||修改系数:"&str_yxs&"→"&str_newxs
		End If
	TmpRs.update
	TmpRs.movenext
	Loop
	TmpRs.close
Call JsAlert(str_lsh&"调试分值系数更改成功！","atask_changexs.asp")
%>
