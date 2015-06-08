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

//=================�ָ���:����js��ȡ���ݿ����=========================
//�����ݿ����ӳ���
function scDBPool(){
    try{
         this.con=new ActiveXObject("ADODB.Connection");
         this.con.Provider="Microsoft.Jet.OLEDB.4.0";
         this.rs=new ActiveXObject("ADODB.Recordset");
    }catch(e){
         this.con=null;
         this.rs=null;
    }
    this.filePath=null;
    this.dbPath=null;
};


//�������ݿ��ļ���ԣ���λ�ļ���·�������ݿ���
scDBPool.prototype.setDB=function(dbPath){
    this.dbPath=dbPath;
};

//�������ݿⶨλ�ļ�,��һ�����Խ��������У�����д�Ƿ���ʹ���κ����ֵ����ݿ�
scDBPool.prototype.setDBPathPosition=function(urlFile){
    var filePath=location.href.substring(0, location.href.indexOf(urlFile));
    this.dbPath=(this.dbPath==null||this.dbPath=="") ? "/Databases/#2012jsb.mdb" : this.dbPath;
    var path=filePath+this.dbPath;
   //ȥ��pathǰ���"files://"�ַ���
    this.filePath=path.substring(8);
};

//ͬ���ݿ⽨������
scDBPool.prototype.connect=function(){
    this.con.ConnectionString="Data Source="+this.filePath;
    this.con.open;
};

//ִ�����ݿ���䷵�ؽ����
scDBPool.prototype.executeQuery=function(sql){
    this.rs.open(sql,this.con);
};

//ִ�����ݿ���䲻���ؽ����
scDBPool.prototype.execute=function(sql){
    this.con.execute(sql);
};

//�رս����
scDBPool.prototype.rsClose=function(){
    this.rs.close();
    this.rs=null;
};

//�ر���������
scDBPool.prototype.conClose=function(){
    this.con.close();
    this.con=null;
};