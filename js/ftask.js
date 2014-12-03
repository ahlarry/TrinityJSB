//�ƻ�������ڵĶ�̬����
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
function checkinf()
{
	fzcheck();
	var objdm=document.frm_ftask
	
	if (objdm.rwlx.value==""){alert("��ѡ����������!"); objdm.rwlx.focus(); return false;}
	else if(objdm.rwlx.value=="��������")
	{
	if (trim(objdm.xldh.value)==""){alert("�����Ų���Ϊ��!\n������������!"); objdm.xldh.value="";objdm.xldh.focus(); return false;}
	
	if(trim(objdm.yhdw.value)==""){alert("�û���λ����Ϊ��!\n�������û���λ!"); objdm.yhdw.value="";objdm.yhdw.focus(); return false;}
	if(trim(objdm.mjmc.value)==""){alert("ģ�����Ʋ���Ϊ��!\n������ģ������!"); objdm.mjmc.value="";objdm.mjmc.focus(); return false;}
	if(trim(objdm.gzyy.value)==""){alert("�������������ԭ����Ϊ��!\n������������������ԭ��!"); objdm.gzyy.value="";objdm.gzyy.focus(); return false;}
	if(trim(objdm.zbfa.value)==""){alert("׼����ȡ��������Ϊ��!\n������׼����ȡ����!"); objdm.zbfa.value="";objdm.zbfa.focus(); return false;}
	if (objdm.zf1.value==0){alert("��ֵ����Ϊ��!"); objdm.zf1.focus(); return false;}
	if (objdm.ed1.value==0){alert("��Ȳ���Ϊ��!"); objdm.ed1.focus(); return false;}
	if (objdm.zrr1.value==""){alert("��ѡ��������!"); objdm.zrr1.focus(); return false;}
	}
	else 
	{
		if (trim(objdm.rwlr.value)==""){alert("�������ݲ���Ϊ��!\n��������������!"); objdm.rwlr.value="";objdm.rwlr.focus(); return false;} 
		if (objdm.zf.value==0){alert("��ֵ����Ϊ��!"); objdm.zf.focus(); return false;}
		if (objdm.ed.value==0){alert("��Ȳ���Ϊ��!"); objdm.ed.focus(); return false;}
		if (objdm.zrr.value==""){alert("��ѡ��������!"); objdm.zrr.focus(); return false;}
	}
	return true;
}

function fzcheck()
{
	var zf=0;
	var ied=0;
	var objdm=document.frm_ftask
	if (objdm.rwlx.value=="��������")
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
  if (name=="��������")
  {
document.all.table1.style.display='';
document.all.table2.style.display='none';  
}
  else{
document.all.table2.style.display='';
document.all.table1.style.display='none';  
}
 }
