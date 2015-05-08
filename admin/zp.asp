<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="upload_del_file.asp" -->
<%
'管理员删除删除文章-----------
if  Request.QueryString("del")<>"" then
 del_image Request.QueryString("del"),"../img/","image","customer"
 delsql="delete from customer where id="&Request.QueryString("del")
 conn.execute(delsql)
end if
'删除结束---------------

'管理员操作开放/停止标志-----------
if  Request.QueryString("stopflag")<>"" then
    sql="update customer set stopflag="&Request.QueryString("stopflag")&" where id="&Request.QueryString("id")
   conn.execute(sql)
end if
'操作开放/停止标志结束---------------%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="Inc/southidc.css" rel="stylesheet" type="text/css">

<script language="JavaScript">
<!--
function cform(){
 if(!confirm("您确认删除吗? 请注意删除后无法恢复!"))
 return false;
}//-->
</script>
</head>

<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center"> 
  <table width="96%" border="0" align="center" class="px14">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td align="center" bgcolor="#6699FF" class="x14"><font color="#FFFFFF">客 户 管 理</font></td>
    </tr>
  </table>
  <table width="96%" border="0">
    <tr> 
      <td valign="top">
        <% sql="select id,sitename,stopflag,datetime from customer order by datetime desc"
	      set rscust=Server.Createobject("ADODB.Recordset")
          rscust.open sql,conn,1,3%> 
        <table width="100%" border="0" cellspacing="1">
          <tr bgcolor="#f0f0f0"> 
            <td width="19%">时间</td>
            <td width="45%">客户名称</td>
            <td width="9%" align="center">状态</td>
            <td width="9%" align="center">操作</td>
            <td width="6%" align="center">修改</td>
            <td width="12%" align="center">删除</td>
          </tr>
          <%do while not rscust.eof%>
          <tr bgcolor="#F6F6F6"> 
            <td><%=rscust("datetime")%> </td>
            <td><a href="zp_modi.asp?id=<%=rscust("id")%>"><%=rscust("sitename")%>&nbsp;</a></td>
            <td align="center"> <% If rscust("stopflag")=0 Then %> <font color="#009900">开放</font> <% Else %> <font color="#FF0000">停止</font> <% End If %> </td>
            <td align="center"> <% If rscust("stopflag")=0 Then %> <a href="?stopflag=1&id=<%=rscust("id")%>">停止</a> 
              <% Else %> <a href="?stopflag=0&id=<%=rscust("id")%>">开放</a> 
              <% End If %> </td>
            <td align="center"><a href="zp_modi.asp?id=<%=rscust("id")%>">修改</a></td>
            <td align="center"><a href="?del=<%=rscust("id")%>"  onclick="return cform()">删除</a></td>
          </tr>
          <%rscust.movenext
		  loop%>
          <%set rscust=nothing%>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>
