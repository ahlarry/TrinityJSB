function tscheckinf()
{
	var objdm=document.frm_mtestadd
	if (objdm.tsyy.value==""){alert("调试原因不能为空!"); objdm.tsyy.focus(); return false;}
	if (objdm.tslr.value==""){alert("调试内容不能为空!"); objdm.tslr.focus(); return false;}
	return true;
}

function tspscheckinf()
{
	var objdm=document.frm_mtestpsadd
	if (objdm.tslr.value==""){alert("评审内容不能为空!"); objdm.tslr.focus(); return false;}
	if (objdm.tsyy.value==""){alert("评审人不能为空!"); objdm.tsyy.focus(); return false;}
	return true;
}