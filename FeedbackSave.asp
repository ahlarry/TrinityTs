<!--#include file="Inc/Conn.asp"-->
<!--#include file="Inc/articleCHAR.INC"-->
<%
If request.form("title")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û������⣬�뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if
If request.form("content")="" Then
Response.Write("<script language=""JavaScript"">alert(""������û�����������ݣ��뷵�ؼ�飡��"");history.go(-1);</script>")
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Feedback"
rs.open sql,conn,1,3

rs.addnew
if session("username")="" then
  rs("Username")="δע���û�"
else
  rs("Username")=trim(request.form("Username"))
end if
rs("Add")=trim(request.form("Add"))
rs("title")=trim(request.form("title"))
rs("content")=htmlencode2(request.form("content"))
rs("Publish")=trim(request.form("Publish"))
rs("time")=date()
rs.update
rs.close

response.redirect "FeedbackView.asp"
%>