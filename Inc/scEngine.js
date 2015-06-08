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