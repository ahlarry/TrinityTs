<!--#include file="conn.asp"-->
<!--#include file="Admin.asp"-->
<!--#INCLUDE file="editor_ubbcode.asp"-->
<%
content=Request.Form("content")
content=ubbcode(replace(Content,strSiteUrl,""))
content=replace(content,"'","""")
'Response.Write content
'Response.end	 

 '������;�ش���
' conn.BeginTrans
   set rs=server.createobject("adodb.recordset")
    if Request("id")="" then
	 rs.open "select * from market",conn,1,3
	 rs.addnew	 
	else
	 rs.open "select * from market where id="&Request("id"),conn,1,3	 
	end if
	 rs("province")=clng(request("province"))
	 rs("title")=request("title")
     rs("content")=content
	 'rs("image")=request("image")
	 'if datetime<>"" then
	 'rs("datetime")=datetime
	 'end if
     rs.update
     rs.Close
    set rs=nothing
'Response.Write request("title")
'Response.end
'  if conn.Errors.Count=0 then
'   conn.CommitTrans 
'  else
'   conn.RollbackTrans 
 ' end if
  '���������;�ش���
  if Request("id")="" then
  Response.Redirect "market.asp?province="&Request("province")
  else
  Response.Write "<script language=javascript>alert('�ɹ���');history.go(-1);</script>"
  end if
  'response.redirect Request.ServerVariables("HTTP_REFERER")

%>