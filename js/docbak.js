function lTrim(str) 
{ 
if (str.charAt(0) == " ") 
{ 
//如果字串左边第一个字符为空格 
str = str.slice(1);//将空格从字串中去掉 
//这一句也可改成 str = str.substring(1, str.length); 
str = lTrim(str); //递归调用 
} 
return str; 
} 

//去掉字串右边的空格 
function rTrim(str) 
{ 
var iLength; 

iLength = str.length; 
if (str.charAt(iLength - 1) == " ") 
{ 
//如果字串右边第一个字符为空格 
str = str.slice(0, iLength - 1);//将空格从字串中去掉 
//这一句也可改成 str = str.substring(0, iLength - 1); 
str = rTrim(str); //递归调用 
} 
return str; 
} 

//去掉字串两边的空格 
function trim(str) 
{ 
return lTrim(rTrim(str)); 
} 


//检测提交信息
function docbak_checkinf()
{
	var objdm=document.frm_docbak
	if (trim(objdm.diskid.value)==""){alert("所存盘号不能为空!"); objdm.diskid.focus(); return false;}
	return true;
}

