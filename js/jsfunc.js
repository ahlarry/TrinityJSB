//��������
function open_win(url,name,width,height,scroll)
{
  //var Left_size = (screen.width>width) ? (screen.width-width)/2 : 0;
  //var Top_size = (screen.height>height) ? (screen.height-height)/2 : 0;
  var Left_size = (screen.width-width)/2;
  var Top_size = (screen.height-height)/2;
  var open_win=window.open(url,name,'width=' + width + ',height=' + height + ',left=' + Left_size + ',top=' + Top_size + ',toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=' + scroll + ',resizable=yes' );
}


function searchlsh_true()
{
  if (document.frm_searchlsh.s_lsh.value=="")
	  { alert("��������Ҫ���ҵ���ˮ�� ��"); document.frm_searchlsh.s_lsh.focus(); return false;}
}
function searchxldh_true()
{
  if (document.frm_searchxldh.s_xldh.value=="")
	  { alert("��������Ҫ���ҵ������� ��"); document.frm_searchlsh.s_xldh.focus(); return false;}
}

/**
*��֤��������
* @��Ҫ��֤�Ŀؼ���ID
* @������������('int', 'u_int', 'float' or 'u_float')
* @���������������ֵ
* @������Ϣ��ʾ�ؼ�ID
* @���ؿ�ֵ
*/
function validationNumber(hndID, numType, maxNum, hndMsgID)
{
 var keyCode = window.event.keyCode;
 var ch = String.fromCharCode(keyCode);
 var value = '';
 var retCode = 0;
 var oNumVerify=null;

 //���λس���
 if (keyCode==13)
 {
  window.event.keyCode = 0;
  return null;
 }

 if (hndID != undefined)
 {
  value += hndID.value+ch;
 }
 else
 {
  window.event.keyCode = keyCode;
  return null;
 }

 oNumVerify = new isNumeric(value);
 if (oNumVerify.isNumber)
 {
  numType = numType.toString().toLowerCase();
  switch(numType)
  {
   case 'u_int':  //������
    if ((oNumVerify.isMinus==false) && (oNumVerify.isDecimal==false))
    {
     retCode = keyCode;
    }
    break;
   case 'u_float':  //��ʵ��
    if (oNumVerify.isMinus==false)
    {
     retCode = keyCode;
    }
    break;
   case 'int':   //����
    if (oNumVerify.isDecimal==false)
    {
     retCode = keyCode;
    }
    break;
   case 'float':  //ʵ��
    retCode = keyCode;
    break;
   default:
    retCode = 0;
  }

  //�ж�����������Ƿ񳬹����õ����ֵ
  if ((maxNum!=undefined) && (maxNum.constructor==Number) && (oNumVerify.value > maxNum))
  {
   retCode = 0;
   if (hndMsgID != undefined)
   {
    hndMsgID.innerHTML = '<SPAN style="color:#EE3333">ֵ���ܴ���'+maxNum+'</SPAN>';
   }
  }
 }

 oNumVerify = null;
 window.event.keyCode = retCode; 
 return null;
} 

function isNumeric(verifyNum)
{
 var re = /^([-]{0,1})([0-9]*)([\.]{0,1})([0-9]*)$/g;
 this.isNumber = false;
 this.isMinus = false;
 this.isDecimal = false;
 this.value = verifyNum;

 verifyNum = verifyNum.toString();

 if (re.test(verifyNum))
 {
  this.isNumber = true;
  re.exec(verifyNum);

  //�ж� '-' ����
  if (RegExp.$1=='-')
  {
   this.isMinus = true;
  }

  //�ж� '.' ����
  if (RegExp.$3=='.')
  {
   this.isDecimal = true;
   verifyNum += '0';
  }

  try
  {
   this.value = parseFloat(verifyNum);
  }
  catch(e)
  {
   this.value = 0;
  }
 }

 return;
}