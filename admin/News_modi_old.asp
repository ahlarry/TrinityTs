
<script language = "JavaScript">
var onecount;
subcat = new Array();
        
onecount=0;

function changelocation(locationid)
    {
    document.myform.SmallClassName.length = 1; 
    var locationid=locationid;
    var i;
    for (i=0;i < onecount; i++)
        {
            if (subcat[i][1] == locationid)
            { 
                document.myform.SmallClassName.options[document.myform.SmallClassName.length] = new Option(subcat[i][0], subcat[i][2]);
            }        
        }
    }    

function AddItem(strFileName){
  document.myform.IncludePic.checked=true;
  document.myform.DefaultPicUrl.value=strFileName;
  document.myform.DefaultPicList.options[document.myform.DefaultPicList.length]=new Option(strFileName,strFileName);
  document.myform.DefaultPicList.selectedIndex+=1;
  if(document.myform.UploadFiles.value==''){
	document.myform.UploadFiles.value=strFileName;
  }
  else{
    document.myform.UploadFiles.value=document.myform.UploadFiles.value+"|"+strFileName;
  }
}
  
function CheckForm()
{
   if (editor.EditMode.checked==true)
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerText;
   else
	  document.myform.Content.value=editor.HtmlEdit.document.body.innerHTML; 

	if (document.myform.title.value.length == 0) {
		alert("新闻标题没有填写.");
		document.myform.title.focus();
		return false; 
	}
		if (document.myform.user.value.length == 0) {
		alert("新闻发布人没有填写");
		document.myform.user.focus();
		return false;
	}
	return true;
}
</script>

<script language="javascript">
<!--  
  if (window == top)top.location.href = "Default.asp"; 
// -->
</script>
<html>
<head>
<title>网软天下企业网站管理系统超漂亮版企业网站管理系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.back_southidc{BACKGROUND-IMAGE:url('image/titlebg.gif');COLOR:000000;}
.table_southidc{BACKGROUND-COLOR: A4B6D7;}
.td_southidc{BACKGROUND-COLOR: F2F8FF;}
.tr_southidc{BACKGROUND-COLOR: ECF5FF;}

.t1 {font:12px 宋体;color:#000000} 
.t2 {font:12px 宋体;color:#ffffff} 
.t3 {font:12px 宋体;color:#ffff00} 
.t4 {font:12px 宋体;color:#800000} 
.t5 {font:12px 宋体;color:#191970} 

.weiqun:hover{
	Font-unline:yes;
	font-family: "宋体";
	color: #FFFFFF;
	text-decoration: underline;
	background-color: #CCCCCC;
}
td {
	font-size: 12px;
}

a:link {
	color: #000000;
	text-decoration: none;
}
a:active {
	color: #000000;
	text-decoration: none;
}
a:visited {
	color: #000000;
	text-decoration: none;
}
.tit {
	font-size: 14px;
}
-->
</style>
<body>

<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top">
      <table  width="95%" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <form method="POST" name="myform" onSubmit="return CheckForm();" action="News_save.asp?action=Modify" target="_self">
          <tr align="center"> 
            <td class="back_southidc" height="30" colspan="2"><font color="#0000FF"><strong>修改新闻</strong></font></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td width="20%" height="24" align="right"><div align="left"><font color="#FF0000">*</font>新闻标题：</div></td>
            <td width="80%" valign="top">　 
              <div align="left">
                <input name="title" type="text" class="input" value="本机调试好系统后，如何上传到空间上？" size="50" maxlength="200">
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left"><font color="#FF0000">*</font>新闻类别：</div></td>
            <td valign="top"> 　 
              <div align="left">
                
             <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                  
                  <option  value="企业新闻">企业新闻</option>
                  
                  <option selected value="业内资讯">业内资讯</option>
                  
            </select>
                <select name="SmallClassName">
                  <option value="" selected>不指定小类</option>
                  
                </select>
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td align="right" valign="top"><div align="left"><font color="#FF0000">*</font>新闻内容：</div></td>
            <td valign="top">
			<textarea name="Content" style="display:none"><P>问：<BR>我在本机调试好企业网站管理系统后，怎么样上传到我的空间上呢？</P>
<P>答：<BR>如果是ACCESS数据库，则非常好办，直接把本机的所有文件（包括数据库）用FTP上传到空间上即可。</P>
<P>如果是SQL数据库，就麻烦一些。SQL版要分程序（及图片）文件的上传和数据库上传两步。程序（及图片）文件的上传非常简单，用FTP直接上传到空间商即可。接下来是数据库的上传。SQL是不能直接用FTP上传到空间上的。你可以使用下面两种方法来做：</P>
<P>1、将本机的SQL数据库进行备份成一个文件，如0791idc.bak，将此文件上传到空间上，然后请空间商帮你还原到你的数据库中。这种方法需要空间商的协助，如果你和空间商关系比较好，可以采用此方法，省事省力。</P>
<P>2、在本机重新安装一个ACCESS版，然后在ACCESS版中运行动易提供的数据迁移/转换程序，把SQL数据库中的内容迁移到ACCESS数据库中，再将这个数据库上传到空间上，在空间上再次运行数据迁移/转换程序，把ACCESS数据库中的数据迁移到SQL数据库中。这种方法简单写就是SQL——ACCESS——SQL。</P>
<P>3、如果是SQL版，本机测试时最好就直接联上远程服务器的SQL数据库，这样虽然会慢一些，但不用担心数据库迁移等问题。本机调试好后，放到空间上时，只要把SQL连接字符串中的主机IP由实际IP改为127.0.0.1即可。</P>
<P>千万不要使用SQL2000自带的导入、导出功能，因为这样做将会带来严重的后果：1、可能会导致数据库中的主键、索引、标识等数据库约束统统丢失。</P></textarea> 
              <iframe ID="editor" src="../editor.asp?Action=News&ID=38" frameborder=1 scrolling=no width="620" height="405"></iframe>
		</td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left">首页图片： 
                <input name="IncludePic" type="hidden" id="IncludePic" value="yes">
              </div></td>
            <td valign="top"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="" size="50" maxlength="200"> 
              <br>
              首页的图片,直接从上传图片中选择： 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option value=""selected>不指定首页图片</option>
                
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles" value=""> 
            </td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left"><font color="#FF0000">*</font>发布人：</div></td>
            <td valign="top">　 
              <div align="left">
                <input name="user" type="text" class="input" size="30" value="admin">
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left">首页图片新闻：</div></td>
            <td>　 
              <div align="left">
                <input type="radio" value="Yes"   name="OK">
                是 
                <input type="radio" value="No" checked name="Ok">
                否 　<font color="#FF0000">（设为首页图片新闻，选择此项时请注意文章中是否添加有图片！）</font></div></td>
          </tr>
          <tr align="center" bgcolor="#ECF5FF"> 
            <td height="35"><div align="left">录入时间：</div></td>
            <td height="35"><div align="left">
                <input name="AddDate" type="text" id="AddDate" value="2005-11-11" maxlength="50">
              </div></td>
          </tr>
          <tr align="center" bgcolor="#ECF5FF"> 
		   <input type="hidden" name="Id" value="38">
            <td height="35" colspan="2"> <input type="submit" name="Submit" value="提交" class="input">          
              　 
              <input type="reset" name="Submit" value="重置" class="input"> </td>
          </tr>
        </form>
      </table>
  </td>
  </tr>
</table>
<table width="100%" height="28" border="0" cellpadding="0" cellspacing="0" class="HeaderTdStyle">
  <tr> 
    <td> 
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="2">
        <tr> 
          <td align="right">Design By <a href="mailto:webmaster@wrtx.cn">heweiqun</a></td>
         
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>