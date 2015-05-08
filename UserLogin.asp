<!--#include file="inc/conn.asp"-->
<!--#include file="inc/md5.asp"-->
<%
dim sql
dim rs
dim username
dim password
username=replace(trim(request("username")),"'","")
password=replace(trim(Request("password")),"'","")
password=md5(password)
set rs=server.createobject("adodb.recordset")
sql="select * from [User] where LockUser=False and username='" & username & "' and password='" & password &"'"
rs.open sql,conn,1,3
if not(rs.bof and rs.eof) then
	if password=rs("password") then
	    rs("LoginIP")=Request.ServerVariables("REMOTE_ADDR")
		rs("LastLoginTime")=now()
		rs("logins")=rs("logins")+1
		rs.update
		session("UserName")=rs("username")		
		session("GZZ")=rs("gzz")		
		Response.Redirect "index.asp"
	end if
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
%>
<html>
<head>
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt; color: #333333; text-decoration: none}
.STYLE1 {color: #FFFFFF}
-->
</style>
</head>
<body>
<br><br><br>
<table width='300' border='0' align='center' cellpadding='4' cellspacing='0' bgcolor="#FFCC00" class='border'>
	<tr >
		
    <td height='15' colspan='2' align="center" bgcolor="#CC6600" class='title STYLE1'>操作: 确认身份失败!</td>
</tr>
<tr>
    <td height='23' colspan='2' align="center" class='tdbg'> <br>
      <br>
      用户名或密码错误！或此帐号被锁定！！<br>
      <br> <a href='javascript:onclick=history.go(-1)'>【返回】</a> <br>
      <br></td>
</tr></table>
</body></html>