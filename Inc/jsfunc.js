/**
*验证数字输入
* @需要验证的控件的ID
* @允许数字类型('int', 'u_int', 'float' or 'u_float')
* @允许输入数字最大值
* @错误消息提示控件ID
* @返回空值
*/
function validationNumber(hndID, numType, maxNum, hndMsgID)
{
 var keyCode = window.event.keyCode;
 var ch = String.fromCharCode(keyCode);
 var value = '';
 var retCode = 0;
 var oNumVerify=null;

 //屏蔽回车键
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
   case 'u_int':  //正整数
    if ((oNumVerify.isMinus==false) && (oNumVerify.isDecimal==false))
    {
     retCode = keyCode;
    }
    break;
   case 'u_float':  //正实数
    if (oNumVerify.isMinus==false)
    {
     retCode = keyCode;
    }
    break;
   case 'int':   //整数
    if (oNumVerify.isDecimal==false)
    {
     retCode = keyCode;
    }
    break;
   case 'float':  //实数
    retCode = keyCode;
    break;
   default:
    retCode = 0;
  }

  //判断输入的数字是否超过设置的最大值
  if ((maxNum!=undefined) && (maxNum.constructor==Number) && (oNumVerify.value > maxNum))
  {
   retCode = 0;
   if (hndMsgID != undefined)
   {
   // hndMsgID.innerHTML = '<SPAN style="color:#EE3333">值不能大于'+maxNum+'</SPAN>';
    alert("值不能大于"+maxNum+"!");
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

  //判断 '-' 符号
  if (RegExp.$1=='-')
  {
   this.isMinus = true;
  }

  //判断 '.' 符号
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

//=================分割线:以下js读取数据库程序=========================
//仿数据库连接池类
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


//设置数据库文件相对（定位文件）路径和数据库名
scDBPool.prototype.setDB=function(dbPath){
    this.dbPath=dbPath;
};

//设置数据库定位文件,这一步可以进连接类中，这里写是方便使用任何名字的数据库
scDBPool.prototype.setDBPathPosition=function(urlFile){
    var filePath=location.href.substring(0, location.href.indexOf(urlFile));
    this.dbPath=(this.dbPath==null||this.dbPath=="") ? "/Databases/#2012jsb.mdb" : this.dbPath;
    var path=filePath+this.dbPath;
   //去除path前面的"files://"字符串
    this.filePath=path.substring(8);
};

//同数据库建立连接
scDBPool.prototype.connect=function(){
    this.con.ConnectionString="Data Source="+this.filePath;
    this.con.open;
};

//执行数据库语句返回结果集
scDBPool.prototype.executeQuery=function(sql){
    this.rs.open(sql,this.con);
};

//执行数据库语句不返回结果集
scDBPool.prototype.execute=function(sql){
    this.con.execute(sql);
};

//关闭结果集
scDBPool.prototype.rsClose=function(){
    this.rs.close();
    this.rs=null;
};

//关闭数据连接
scDBPool.prototype.conClose=function(){
    this.con.close();
    this.con=null;
};