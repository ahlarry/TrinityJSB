function tscheckinf()
{
	var objdm=document.frm_mtestadd
	if (objdm.tsyy.value==""){alert("����ԭ����Ϊ��!"); objdm.tsyy.focus(); return false;}
	if (objdm.tslr.value==""){alert("�������ݲ���Ϊ��!"); objdm.tslr.focus(); return false;}
	return true;
}

function tspscheckinf()
{
	var objdm=document.frm_mtestpsadd
	if (objdm.tslr.value==""){alert("�������ݲ���Ϊ��!"); objdm.tslr.focus(); return false;}
	if (objdm.tsyy.value==""){alert("�����˲���Ϊ��!"); objdm.tsyy.focus(); return false;}
	return true;
}