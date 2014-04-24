//�ƻ�������ڵĶ�̬����
//15:18 2007-1-6-������
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
	var objdm=document.mtask_add
	if (objdm.ddh.value==""){alert("�����Ų���Ϊ��!"); objdm.ddh.focus(); return false;}
	if (objdm.lsh.value==""){alert("��ˮ�Ų���Ϊ��!"); objdm.lsh.focus(); return false;}
	if (objdm.dwmc.value==""){alert("�ͻ����Ʋ���Ϊ��!"); objdm.dwmc.focus(); return false;}
	if (objdm.dmmc.value==""){alert("�������Ʋ���Ϊ��!"); objdm.dmmc.focus(); return false;}
	if (objdm.mh.value==""){alert("ģ�Ų���Ϊ��!"); objdm.mh.focus(); return false;}
	if (objdm.sbcj.value==""){alert("�豸���Ҳ���Ϊ��!"); objdm.sbcj.focus(); return false;}
	if (objdm.jcjxh.value==""){alert("�������ͺŲ���Ϊ��!"); objdm.jcjxh.focus(); return false;}
	if (objdm.rdogg.value==""){alert("�ȵ�ż�����Ϊ��!"); objdm.rdogg.focus(); return false;}
	if (objdm.mjcl.value==""){alert("ģ�߲��ϲ���Ϊ��!"); objdm.mjcl.focus(); return false;}
	if (objdm.mjzf.value==0){alert("ģ�߷�ֵ���� 0 !"); objdm.ckdm.focus(); return false;}
	if (objdm.mtbl.value==""){alert("ģͷ�������ܿ�!"); objdm.mtbl.focus(); return false;}
	if (objdm.dxbl.value==""){alert("���ͱ������ܿ� !"); objdm.dxbl.focus(); return false;}
	if (objdm.mtjgbl.value==""){alert("ģͷ�ṹ�������ܿ� !"); objdm.mtjgbl.focus(); return false;}
	if (objdm.dxjgbl.value==""){alert("���ͽṹ�������ܿ� !"); objdm.dxjgbl.focus(); return false;}
	if (objdm.fzxs.value==0){alert("����ϵ������ 0 !"); objdm.fzxs.focus(); return false;}
	if (objdm.tsdzf.value==0){alert("�������������ʱ������ѡ���ֵ !"); objdm.ckdm.focus(); return false;}
	if (objdm.qysd.value==""){alert("ǣ���ٶȲ���Ϊ��!"); objdm.qysd.focus(); return false;}
	if (objdm.cnts.value=="true") and (objdm.tlsb.value==""){alert("���ڵ���ʱ���������Ϊ��!"); objdm.tslb.focus(); return false;}
	if (objdm.xcbh.value==""){alert("�Ͳıں���Ϊ��!"); objdm.xcbh.focus(); return false;}
	if (objdm.mjxx.value!="ģͷ" && objdm.dxjg.value==""){alert("���ͽṹ����Ϊ��!"); objdm.dxjg.focus(); return false;}
	if (objdm.mjxx.value!="ģͷ" && objdm.sxjg.value==""){alert("ˮ��ṹ����Ϊ��!"); objdm.sxjg.focus(); return false;}
	if (objdm.mtljcc.value==""){alert("ģͷ���ӳߴ粻��Ϊ��!"); objdm.mtljcc.focus(); return false;}
	if (objdm.mtbl.value==""){alert("ģͷ��������Ϊ��!"); objdm.mtbl.focus(); return false;}
	if (objdm.jgzz.value==""){alert("��ѡ��ṹ�鳤"); objdm.jgzz.focus(); return false;}
	if (objdm.sjzz.value==""){alert("��ѡ������鳤"); objdm.sjzz.focus(); return false;}
	if (objdm.jsdb.value==""){alert("��ѡ��������!"); objdm.jsdb.focus(); return false;}
	return true;
}

