function lTrim(str) 
{ 
if (str.charAt(0) == " ") 
{ 
//����ִ���ߵ�һ���ַ�Ϊ�ո� 
str = str.slice(1);//���ո���ִ���ȥ�� 
//��һ��Ҳ�ɸĳ� str = str.substring(1, str.length); 
str = lTrim(str); //�ݹ���� 
} 
return str; 
} 

//ȥ���ִ��ұߵĿո� 
function rTrim(str) 
{ 
var iLength; 

iLength = str.length; 
if (str.charAt(iLength - 1) == " ") 
{ 
//����ִ��ұߵ�һ���ַ�Ϊ�ո� 
str = str.slice(0, iLength - 1);//���ո���ִ���ȥ�� 
//��һ��Ҳ�ɸĳ� str = str.substring(0, iLength - 1); 
str = rTrim(str); //�ݹ���� 
} 
return str; 
} 

//ȥ���ִ����ߵĿո� 
function trim(str) 
{ 
return lTrim(rTrim(str)); 
} 

function chkygkp(frm)
{
	var objdm=frm;
	if (trim(objdm.kpzrr.value)==""){alert("��ѡ������Ա!"); objdm.kpzrr.value=""; return false;}
	if (trim(objdm.kpinfo.value)==""){alert("��ѡ������Ŀ!");objdm.kpinfo.value="";objdm.kpinfo.focus();return false;}
	if (trim(objdm.kpbz.value)==""){alert("���µ�֤�ݰ�!�Ժ󱸲�!\n\n�����뱸ע!");objdm.kpbz.value="";objdm.kpbz.focus();return false;}
	return true;
}

function chkpgbkp(frm)
{
	var objdm=frm;
	if (trim(objdm.kplsh.value)==""){alert("���������ģ�ߵ���ˮ��!"); objdm.kplsh.value="";objdm.kplsh.focus(); return false;}
	if (trim(objdm.kpinfo.value)==""){alert("��ѡ������Ŀ!");objdm.kpinfo.value="";objdm.kpinfo.focus();return false;}
	if (trim(objdm.kpbz.value)==""){alert("���µ�֤�ݰ�!�Ժ󱸲�!\n\n�����뱸ע!");objdm.kpbz.value="";objdm.kpbz.focus();return false;}
	return true;
}