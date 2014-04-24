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

function chkygkp(frm)
{
	var objdm=frm;
	if (trim(objdm.kpzrr.value)==""){alert("请选择考评人员!"); objdm.kpzrr.value=""; return false;}
	if (trim(objdm.kpinfo.value)==""){alert("请选择考评项目!");objdm.kpinfo.value="";objdm.kpinfo.focus();return false;}
	if (trim(objdm.kpbz.value)==""){alert("留下点证据吧!以后备查!\n\n请输入备注!");objdm.kpbz.value="";objdm.kpbz.focus();return false;}
	return true;
}

function chkpgbkp(frm)
{
	var objdm=frm;
	if (trim(objdm.kplsh.value)==""){alert("请输入相关模具的流水号!"); objdm.kplsh.value="";objdm.kplsh.focus(); return false;}
	if (trim(objdm.kpinfo.value)==""){alert("请选择考评项目!");objdm.kpinfo.value="";objdm.kpinfo.focus();return false;}
	if (trim(objdm.kpbz.value)==""){alert("留下点证据吧!以后备查!\n\n请输入备注!");objdm.kpbz.value="";objdm.kpbz.focus();return false;}
	return true;
}