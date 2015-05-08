<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Inc/southidc.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
  <table width="96%" border="0" align="center" class="px14">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#6699FF" class="x14"><font color="#FFFFFF">修 改 客 户 </font></td>
    </tr>
  </table>
  <table width="96%" border="0">
    <tr> 
      <td valign="top"> 
        
        <script>
<!--
function validata_form() {
validity=true;

if (!check_empty(document.form2.classid.value))
{ validity=false;
 alert('请选择类别');
document.form2.classid.focus();
 return validity; }
 

if (!check_empty(document.form2.sitename.value))
{ validity=false;
 alert('请填写网站名');
document.form2.sitename.focus();
 return validity; }
if (!check_empty(document.form2.content.value))
{ validity=false;
 alert('请填写介绍内容');
document.form2.content.focus();
 return validity; } 
 
function check_empty(text) {
return (text.length > 0); // returns false if empty
}
}
-->
</script> 
        <table width="100%" border="0" cellspacing="1">
          <form action="zp_write.asp" method="post" name="form2" onsubmit="return validata_form()">
		   <%id=Request.QueryString("id")
		 sql="select * from customer where id="&id
	      set rs=Server.Createobject("ADODB.Recordset")
         rs.open sql,conn,1,3		 
		%>
            <tr bgcolor="#f0f0f0"> 
              <td width="11%" align="right">类别:</td>
              <td width="89%"> <% sql="select * from customer_class"
	               set rscust=Server.Createobject("ADODB.Recordset")
                   rscust.open sql,conn,1,3
				  %> <select name="classid" class="input">
                  <option value="">请选择类别</option>
                  <%do while not rscust.eof%>
                  <option value="<%=rscust("id")%>" <% If rscust("id")=rs("classid") Then %>selected<% End If %>><%=rscust("classname")%></option>
                  <%rscust.movenext
				   loop%>
                  <%set rscust=nothing%>
                </select>                <a href="class_customer.asp">增加类别</a>
              </td>
            </tr>
            <tr bgcolor="#f0f0f0"> 
              <td align="right">客户名称:</td>
              <td><input name="sitename" type="text" class="input" id="sitename" size="60" maxlength="100"  value="<%=rs("sitename")%>">
                (100字符以内)</td>
            </tr>
            <tr bgcolor="#f0f0f0"> 
              <td align="right">网 址:</td>
              <td><input name="url" type="text" class="input" id="url" size="60" maxlength="50" value="<%=rs("url")%>"></td>
            </tr>
            <tr bgcolor="#f0f0f0"> 
              <td align="right">图 标:</td>
              <td> <input name="image" type="text" class="input" id="image" size="40"  maxlength="40"  value="<%=rs("image")%>"> 
                <input type="button" name="Submit" value="上传" onClick="window.open('upload_flash.asp?formname=form2&editname=image&uppath=UploadFiles&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')">
                70x35像素，
                <% If rs("image")="" Then %>
                如没有图片, 就不要填写 
                <% Else %>
                <input type="button" name="Submit4" value="要更换图片,请先删除当前图片" onClick="window.open('upload_delete_file.asp?formname=form2&editname=image&i=<%=rs("id")%>&p=../img/&f=image&t=zuopin','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')"> 
                <% End If %></td>
            </tr>
            <tr bgcolor="#f0f0f0">
              <td align="right">主要客户:</td>
              <td><input name="ismain" type="checkbox" id="ismain" value="checkbox" <%if rs("ismain")=True then Response.Write("checked") %>>
              (显示在客户首页) </td>
            </tr>
            <tr bgcolor="#f0f0f0"> 
              <td align="right"><input type="hidden" name="id" value="<%=rs("id")%>"></td>
              <td><input type="submit" name="Submit2" value="修改"> <input type="reset" name="Submit3" value="重写"></td>
            </tr>
          </form>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>
