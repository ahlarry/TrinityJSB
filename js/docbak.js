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


//����ύ��Ϣ
function docbak_checkinf()
{
	var objdm=document.frm_docbak
	if (trim(objdm.diskid.value)==""){alert("�����̺Ų���Ϊ��!"); objdm.diskid.focus(); return false;}
	return true;
}

