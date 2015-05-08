<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
'	Call ChkPageAble(0)
	Dim iid,strlsh
	iid=Request("id")
	strlsh=Request("lsh")
	If iid="" Or Not isNumeric(iid) Then 
		errmsg="请从正确入口进入！"
		Call WriteErrMsg()
	else
		Call kp_delete()	
	End If
		
	Function kp_delete()
		if iid<>"" then
			sql="delete from mission where ID=" & Clng(iid)
			conn.Execute sql
		end if
		if strlsh<>"" then
			sql="delete from ins_tsxx where lsh='"&strlsh&"'"
			conn.Execute sql		
			sql="delete from ins_tsfz where lsh='"&strlsh&"'"
			conn.Execute sql
			sql="delete from ext_tsxx where lsh='"&strlsh&"'"
			conn.Execute sql		
			sql="delete from ext_tsfz where lsh='"&strlsh&"'"
			conn.Execute sql			
		end if
		call CloseConn()      		
		Call WriteSuccessMsg("删除成功","")
	End Function
%>