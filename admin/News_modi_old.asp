
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
		alert("���ű���û����д.");
		document.myform.title.focus();
		return false; 
	}
		if (document.myform.user.value.length == 0) {
		alert("���ŷ�����û����д");
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
<title>����������ҵ��վ����ϵͳ��Ư������ҵ��վ����ϵͳ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.back_southidc{BACKGROUND-IMAGE:url('image/titlebg.gif');COLOR:000000;}
.table_southidc{BACKGROUND-COLOR: A4B6D7;}
.td_southidc{BACKGROUND-COLOR: F2F8FF;}
.tr_southidc{BACKGROUND-COLOR: ECF5FF;}

.t1 {font:12px ����;color:#000000} 
.t2 {font:12px ����;color:#ffffff} 
.t3 {font:12px ����;color:#ffff00} 
.t4 {font:12px ����;color:#800000} 
.t5 {font:12px ����;color:#191970} 

.weiqun:hover{
	Font-unline:yes;
	font-family: "����";
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
            <td class="back_southidc" height="30" colspan="2"><font color="#0000FF"><strong>�޸�����</strong></font></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td width="20%" height="24" align="right"><div align="left"><font color="#FF0000">*</font>���ű��⣺</div></td>
            <td width="80%" valign="top">�� 
              <div align="left">
                <input name="title" type="text" class="input" value="�������Ժ�ϵͳ������ϴ����ռ��ϣ�" size="50" maxlength="200">
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left"><font color="#FF0000">*</font>�������</div></td>
            <td valign="top"> �� 
              <div align="left">
                
             <select name="BigClassName" onChange="changelocation(document.myform.BigClassName.options[document.myform.BigClassName.selectedIndex].value)" size="1">
                  
                  <option  value="��ҵ����">��ҵ����</option>
                  
                  <option selected value="ҵ����Ѷ">ҵ����Ѷ</option>
                  
            </select>
                <select name="SmallClassName">
                  <option value="" selected>��ָ��С��</option>
                  
                </select>
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td align="right" valign="top"><div align="left"><font color="#FF0000">*</font>�������ݣ�</div></td>
            <td valign="top">
			<textarea name="Content" style="display:none"><P>�ʣ�<BR>���ڱ������Ժ���ҵ��վ����ϵͳ����ô���ϴ����ҵĿռ����أ�</P>
<P>��<BR>�����ACCESS���ݿ⣬��ǳ��ð죬ֱ�Ӱѱ����������ļ����������ݿ⣩��FTP�ϴ����ռ��ϼ��ɡ�</P>
<P>�����SQL���ݿ⣬���鷳һЩ��SQL��Ҫ�ֳ��򣨼�ͼƬ���ļ����ϴ������ݿ��ϴ����������򣨼�ͼƬ���ļ����ϴ��ǳ��򵥣���FTPֱ���ϴ����ռ��̼��ɡ������������ݿ���ϴ���SQL�ǲ���ֱ����FTP�ϴ����ռ��ϵġ������ʹ���������ַ���������</P>
<P>1����������SQL���ݿ���б��ݳ�һ���ļ�����0791idc.bak�������ļ��ϴ����ռ��ϣ�Ȼ����ռ��̰��㻹ԭ��������ݿ��С����ַ�����Ҫ�ռ��̵�Э���������Ϳռ��̹�ϵ�ȽϺã����Բ��ô˷�����ʡ��ʡ����</P>
<P>2���ڱ������°�װһ��ACCESS�棬Ȼ����ACCESS�������ж����ṩ������Ǩ��/ת�����򣬰�SQL���ݿ��е�����Ǩ�Ƶ�ACCESS���ݿ��У��ٽ�������ݿ��ϴ����ռ��ϣ��ڿռ����ٴ���������Ǩ��/ת�����򣬰�ACCESS���ݿ��е�����Ǩ�Ƶ�SQL���ݿ��С����ַ�����д����SQL����ACCESS����SQL��</P>
<P>3�������SQL�棬��������ʱ��þ�ֱ������Զ�̷�������SQL���ݿ⣬������Ȼ����һЩ�������õ������ݿ�Ǩ�Ƶ����⡣�������Ժú󣬷ŵ��ռ���ʱ��ֻҪ��SQL�����ַ����е�����IP��ʵ��IP��Ϊ127.0.0.1���ɡ�</P>
<P>ǧ��Ҫʹ��SQL2000�Դ��ĵ��롢�������ܣ���Ϊ����������������صĺ����1�����ܻᵼ�����ݿ��е���������������ʶ�����ݿ�Լ��ͳͳ��ʧ��</P></textarea> 
              <iframe ID="editor" src="../editor.asp?Action=News&ID=38" frameborder=1 scrolling=no width="620" height="405"></iframe>
		</td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left">��ҳͼƬ�� 
                <input name="IncludePic" type="hidden" id="IncludePic" value="yes">
              </div></td>
            <td valign="top"><input name="DefaultPicUrl" type="text" id="DefaultPicUrl" value="" size="50" maxlength="200"> 
              <br>
              ��ҳ��ͼƬ,ֱ�Ӵ��ϴ�ͼƬ��ѡ�� 
              <select name="DefaultPicList" id="DefaultPicList" onChange="DefaultPicUrl.value=this.value;">
                <option value=""selected>��ָ����ҳͼƬ</option>
                
              </select> <input name="UploadFiles" type="hidden" id="UploadFiles" value=""> 
            </td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left"><font color="#FF0000">*</font>�����ˣ�</div></td>
            <td valign="top">�� 
              <div align="left">
                <input name="user" type="text" class="input" size="30" value="admin">
              </div></td>
          </tr>
          <tr bgcolor="#ECF5FF"> 
            <td height="24" align="right"><div align="left">��ҳͼƬ���ţ�</div></td>
            <td>�� 
              <div align="left">
                <input type="radio" value="Yes"   name="OK">
                �� 
                <input type="radio" value="No" checked name="Ok">
                �� ��<font color="#FF0000">����Ϊ��ҳͼƬ���ţ�ѡ�����ʱ��ע���������Ƿ�������ͼƬ����</font></div></td>
          </tr>
          <tr align="center" bgcolor="#ECF5FF"> 
            <td height="35"><div align="left">¼��ʱ�䣺</div></td>
            <td height="35"><div align="left">
                <input name="AddDate" type="text" id="AddDate" value="2005-11-11" maxlength="50">
              </div></td>
          </tr>
          <tr align="center" bgcolor="#ECF5FF"> 
		   <input type="hidden" name="Id" value="38">
            <td height="35" colspan="2"> <input type="submit" name="Submit" value="�ύ" class="input">          
              �� 
              <input type="reset" name="Submit" value="����" class="input"> </td>
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