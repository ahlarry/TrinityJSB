//计划完成日期的动态调整
//15:18 2007-1-6-星期六
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
	var objdm=document.mtask_add
	if (objdm.ddh.value==""){alert("订单号不能为空!"); objdm.ddh.focus(); return false;}
	if (objdm.lsh.value==""){alert("流水号不能为空!"); objdm.lsh.focus(); return false;}
	if (objdm.dwmc.value==""){alert("客户名称不能为空!"); objdm.dwmc.focus(); return false;}
	if (objdm.dmmc.value==""){alert("断面名称不能为空!"); objdm.dmmc.focus(); return false;}
	if (objdm.mh.value==""){alert("模号不能为空!"); objdm.mh.focus(); return false;}
	if (objdm.sbcj.value==""){alert("设备厂家不能为空!"); objdm.sbcj.focus(); return false;}
	if (objdm.jcjxh.value==""){alert("挤出机型号不能为空!"); objdm.jcjxh.focus(); return false;}
	if (objdm.rdogg.value==""){alert("热电偶规格不能为空!"); objdm.rdogg.focus(); return false;}
	if (objdm.mjcl.value==""){alert("模具材料不能为空!"); objdm.mjcl.focus(); return false;}
	if (objdm.mjzf.value==0){alert("模具分值不能 0 !"); objdm.ckdm.focus(); return false;}
	if (objdm.mtbl.value==""){alert("模头比例不能空!"); objdm.mtbl.focus(); return false;}
	if (objdm.dxbl.value==""){alert("定型比例不能空 !"); objdm.dxbl.focus(); return false;}
	if (objdm.mtjgbl.value==""){alert("模头结构比例不能空 !"); objdm.mtjgbl.focus(); return false;}
	if (objdm.dxjgbl.value==""){alert("定型结构比例不能空 !"); objdm.dxjgbl.focus(); return false;}
	if (objdm.fzxs.value==0){alert("复杂系数不能 0 !"); objdm.fzxs.focus(); return false;}
	if (objdm.tsdzf.value==0){alert("更改添加任务书时需重新选择分值 !"); objdm.ckdm.focus(); return false;}
	if (objdm.qysd.value==""){alert("牵引速度不能为空!"); objdm.qysd.focus(); return false;}
	if (objdm.cnts.value=="true") and (objdm.tlsb.value==""){alert("厂内调试时调试类别不能为空!"); objdm.tslb.focus(); return false;}
	if (objdm.xcbh.value==""){alert("型材壁厚不能为空!"); objdm.xcbh.focus(); return false;}
	if (objdm.mjxx.value!="模头" && objdm.dxjg.value==""){alert("定型结构不能为空!"); objdm.dxjg.focus(); return false;}
	if (objdm.mjxx.value!="模头" && objdm.sxjg.value==""){alert("水箱结构不能为空!"); objdm.sxjg.focus(); return false;}
	if (objdm.mtljcc.value==""){alert("模头连接尺寸不能为空!"); objdm.mtljcc.focus(); return false;}
	if (objdm.mtbl.value==""){alert("模头比例不能为空!"); objdm.mtbl.focus(); return false;}
	if (objdm.jgzz.value==""){alert("请选择结构组长"); objdm.jgzz.focus(); return false;}
	if (objdm.sjzz.value==""){alert("请选择设计组长"); objdm.sjzz.focus(); return false;}
	if (objdm.jsdb.value==""){alert("请选择技术代表!"); objdm.jsdb.focus(); return false;}
	return true;
}

