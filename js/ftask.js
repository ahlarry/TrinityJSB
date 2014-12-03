//计划完成日期的动态调整
var normal_day = [31,28,31,30,31,30,31,31,30,31,30,31];
var leap_day = [31,29,31,30,31,30,31,31,30,31,30,31];

function isLeapYear(yyear)
{
	if(yyear%4==0 || (yyear%400==0 && yyear%100!=0))
		return true;
	else
		return false;
}

function removeOptions(optionMenu)
{
for(i=0; i<optionMenu.options.length; i++)
	optionMenu.options[i] = null;
}

function addOptions(yyear,mmonth,optionList)
{
var i = 0;
//var ddd;
removeOptions(optionList);

if(isLeapYear(yyear))
{
for(i=0; i<leap_day[mmonth]; i++)
optionList[i] = new Option(i+1,i+1);
}
else
{
for(i=0; i<normal_day[mmonth]; i++)
optionList[i] = new Option(i+1,i+1);
}
}


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
function checkinf()
{
	fzcheck();
	var objdm=document.frm_ftask
	
	if (objdm.rwlx.value==""){alert("请选择任务类型!"); objdm.rwlx.focus(); return false;}
	else if(objdm.rwlx.value=="零星修理")
	{
	if (trim(objdm.xldh.value)==""){alert("修理单号不能为空!\n请输入修理单号!"); objdm.xldh.value="";objdm.xldh.focus(); return false;}
	
	if(trim(objdm.yhdw.value)==""){alert("用户单位不能为空!\n请输入用户单位!"); objdm.yhdw.value="";objdm.yhdw.focus(); return false;}
	if(trim(objdm.mjmc.value)==""){alert("模具名称不能为空!\n请输入模具名称!"); objdm.mjmc.value="";objdm.mjmc.focus(); return false;}
	if(trim(objdm.gzyy.value)==""){alert("故障现象与分析原因不能为空!\n请输入故障现象与分析原因!"); objdm.gzyy.value="";objdm.gzyy.focus(); return false;}
	if(trim(objdm.zbfa.value)==""){alert("准备采取方案不能为空!\n请输入准备采取方案!"); objdm.zbfa.value="";objdm.zbfa.focus(); return false;}
	if (objdm.zf1.value==0){alert("分值不能为零!"); objdm.zf1.focus(); return false;}
	if (objdm.ed1.value==0){alert("额度不能为零!"); objdm.ed1.focus(); return false;}
	if (objdm.zrr1.value==""){alert("请选择责任人!"); objdm.zrr1.focus(); return false;}
	}
	else 
	{
		if (trim(objdm.rwlr.value)==""){alert("任务能容不能为空!\n请输入任务内容!"); objdm.rwlr.value="";objdm.rwlr.focus(); return false;} 
		if (objdm.zf.value==0){alert("分值不能为零!"); objdm.zf.focus(); return false;}
		if (objdm.ed.value==0){alert("额度不能为零!"); objdm.ed.focus(); return false;}
		if (objdm.zrr.value==""){alert("请选择责任人!"); objdm.zrr.focus(); return false;}
	}
	return true;
}

function fzcheck()
{
	var zf=0;
	var ied=0;
	var objdm=document.frm_ftask
	if (objdm.rwlx.value=="零星修理")
 { 
 	if(!isNaN(parseFloat(document.all.zf1.value)))
 	{
 		zf=parseFloat(document.all.zf1.value);
	  document.all.zf1.value=zf;
    return;
    }
 	if(!isNaN(parseFloat(document.all.ed1.value)))
 	{
 		ied=parseFloat(document.all.ed1.value);
	  document.all.ed1.value=ied;
    return;
    }
  }
  else
  	if(!isNaN(parseFloat(document.all.zf.value)))
 	{
 		zf=parseFloat(document.all.zf.value);
	  document.all.zf.value=zf;
    return;
  	if(!isNaN(parseFloat(document.all.ed.value)))
 	{
 		ied=parseFloat(document.all.ed.value);
	  document.all.ed.value=ied;
    return;
    }
}
function selecttask(s) 
{
  var name = s.options[s.selectedIndex].value;
  if (name=="零星修理")
  {
document.all.table1.style.display='';
document.all.table2.style.display='none';  
}
  else{
document.all.table2.style.display='';
document.all.table1.style.display='none';  
}
 }
