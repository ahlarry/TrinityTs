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