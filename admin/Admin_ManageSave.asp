<%@ language="VBScript"%>
<%response.Expires = 0%>
<!--#include file="Conn.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/function.asp"-->
<%
password=replace(trim(Request("password")),"'","")
UserName=replace(trim(Request("UserName")),"'","")

if strLength(password)>16 or strLength(password)<5 then
 response.write  "<script language=javascript>alert('�����������λ������С��6λ�����16λ!');history.go(-1);</script>"
  response.End
end if

'�������ݿ⣬���˹���Ա�Ƿ��Ѿ�����
set rs=server.createobject("adodb.recordset")
sql="select * from admin where UserName='" & UserName & "'"
rs.open sql,conn,1,1

if rs.recordcount >= 1 then
  response.write  "<script language=javascript>alert('�˹���Ա�ʺ��Ѿ����ڣ���ѡ�������ʺ�!');history.go(-1);</script>"
  response.End
  rs.close
  set rs=nothing
end if

password=md5(password)
set rs=server.createobject("adodb.recordset")
sql="select * from admin"
rs.open sql,conn,1,3

'����һ������Ա�ʺŵ����ݿ�
rs.addnew
rs("UserName")=UserName
rs("PassWord")=password
rs.update
rs.close
set rs=nothing
Response.Redirect "Admin_Manage.asp"
%>