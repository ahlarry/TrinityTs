<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%
 '事务处理和卷回处理
' conn.BeginTrans
   set rs=server.createobject("adodb.recordset")
   if Request("id")="" then
     rs.open "select * from customer",conn,1,3
     rs.addnew
	else
	rs.open "select * from customer where id="&Request("id"),conn,1,3
	end if
     rs("classid")=cint(Request.Form("classid"))
     rs("sitename")=Trim(Request.Form("sitename"))
	 rs("url")=Trim(Request.Form("url"))
     'rs("content")=Trim(Request.Form("content"))
     rs("image")=Trim(Request.Form("image"))
	 if Request.Form("ismain")<>"" then
	  rs("ismain")=True
	 else
	  rs("ismain")=False
	 end if
     rs.update
     rs.Close
    set rs=nothing
'  if conn.Errors.Count=0 then
'   conn.CommitTrans 
'  else
'   conn.RollbackTrans 
 ' end if
  '完成事务处理和卷回处理
if Request("id")="" then
 %> 
   <script language="javascript"> 
  if(confirm("添加成功, 继续返回添加客户吗?\n按确定继续添加，按取消转到客户管理列表。"))
  history.go(-1);
  else
  document.location.href="zp.asp";
  </script> 
 <%  
 else
 Response.Write "<script language=javascript>alert('修改成功！');document.location.href='zp_modi.asp';</script>"
 end if
%>