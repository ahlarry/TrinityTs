<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim stryhm, strxs, strgzz, stract,strgxid,strjg
	stryhm=Trim(Request("yhm")) : strxs=Trim(Request("khxs")) : strgzz=Trim(Request("gzz"))
	stract=Trim(Request("act")) : strgxid=Trim(Request("gengx")) : strjg=0
	
	Select Case stract
		Case "new"
    		'检测流水号是否已存在
    		strSql="select UserName from [user] where UserName='"&stryhm&"'"
    		set rs=server.createobject("adodb.recordset")
    		rs.open strSql,conn,1,1
    		If Not(Rs.eof Or Rs.bof) Then
    			errmsg="用户名 " & stryhm & " 已存在!请更改用户名!"
    			Call WriteErrMsg()
    		Else
        		strSql="select * from [user]"
        		set Rs=server.createobject("adodb.recordset")
        		Rs.open strSql,Conn,1,3
        		Rs.AddNew
        			Rs("UserName")=stryhm
        			Rs("Grxs")=strxs
        			Rs("Gzz")=strgzz
        		Rs.update
        		Rs.Close
        		strjg=1		
    		End If
    		Rs.Close
   		Case "renewal"
  			strgxid=split(strgxid,";")
   			For i=0 to ubound(strgxid)
'   				errmsg=errmsg&"strgxid="&strgxid(i)&",xs="&Request(strgxid(i))&"id="&Replace(strgxid(i), "xs", "") &"<br>"
        		strSql="select * from [user] where UserID=" & Clng(Replace(strgxid(i), "xs", "") )
        		set Rs=server.createobject("adodb.recordset")
        		Rs.open strSql,Conn,1,3
        		If not Rs.eof Then 
        			Rs("Grxs")=Trim(Request(strgxid(i)))
        		End If 
        		Rs.update
        		Rs.Close   				
   			Next
        	If errmsg <> "" Then
        		Call WriteErrMsg()
        	Else
        		strjg=1
        	End If			
		Case Else
        	if stract<>"" then
        		sql="delete from User where UserID=" & Clng(stract)
        		conn.Execute sql
        		strjg=1
        	end if
        	Call CloseConn()
	End Select	
	
	If strjg=1 Then	Call WriteSuccessMsg("用户数据更新成功!","")
%>
