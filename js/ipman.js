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

function chkipman1(frm)
{
	var objdm=frm;
	if (trim(objdm.ip.value)==""){alert("������IP��ַ!"); objdm.ip.value="";objdm.ip.focus(); return false;}
	return true;
}
function chkipman2(frm)
{
	var objdm=frm;
	if (trim(objdm.mac.value)==""){alert("������MAC��ַ!"); objdm.mac.value="";objdm.mac.focus(); return false;}
	return true;
}
function chkipman3(frm)
{
	var objdm=frm;
	if (trim(objdm.ip.value)==""){alert("������IP��ַ!"); objdm.ip.value="";objdm.ip.focus(); return false;}
	return true;
}
function chkipman4(frm)
{
	var objdm=frm;
	if (trim(objdm.ip.value)==""){alert("������IP��ַ!"); objdm.ip.value="";objdm.ip.focus(); return false;}
	return true;
}