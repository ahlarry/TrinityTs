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
   // hndMsgID.innerHTML = '<SPAN style="color:#EE3333">ֵ���ܴ���'+maxNum+'</SPAN>';
    alert("ֵ���ܴ���"+maxNum+"!");
    hndID.value="";
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