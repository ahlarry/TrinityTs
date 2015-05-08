<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim action
	action=Request("action")

	Dim strzrr, strkhm, strlsh, strsj, strzywt, stryyfx, strjzcs
	strzrr=Request("zrr")
	strkhm=Request("khm")
	strlsh=Request("lsh")
	strsj=Request("sj")	
	strzywt=Request("zywt")
	stryyfx=Request("yyfx")
	strjzcs=Request("jzcs")
	errmsg=""
	
	Select case action
		Case "add"
			Call Faq_add()
	End Select
		
Function Faq_add()
	Dim tmpzrr
	If strkhm="" Then errmsg=errmsg & "客户名不能为空!<br>"
   	If strlsh="" Then errmsg=errmsg & "流水号不能为空!<br>"
   	If strsj="" Then errmsg=errmsg & "时间不能为空!<br>"    	
   	If strzywt="" Then errmsg=errmsg & "主要问题不能为空!<br>"    	
   	If strzrr="" Then 
  		errmsg=errmsg & "责任人不能为空!<br>"
   	else 
     	tmpzrr=Replace(strzrr,"，",",")
       	tmpzrr=split(tmpzrr,",")	
   	End If
    '检测用户名是否正确
   	For i=0 to ubound(tmpzrr)
   		strSql="select * from [user] where UserName='"&tmpzrr(i)&"'"
   		set Rs=server.createobject("adodb.recordset")
   		Rs.open strSql,Conn,1,1
    	If rs.eof Then
   			errmsg=errmsg &"调试员 <b>"& tmpzrr(i) &"</b> 不存在，请核对后重新输入!<br>"
    	End If
    	Rs.Close
   	Next

   	If errmsg <> "" Then
   		Call WriteErrMsg()
   		Response.End
    End If
    	
   	'添加问题分析
   	strSql="select * from [Faq]"
   	set Rs=server.createobject("adodb.recordset")
    Rs.open strSql,Conn,1,3
       	Rs.AddNew
       	Rs("zrr")=strzrr
       	Rs("Khmc")=strkhm
       	Rs("Lsh")=strlsh
       	Rs("Sj")=strsj
       	Rs("WenTi")=strzywt
       	If stryyfx<>"" Then Rs("YuanY")=stryyfx
       	If strjzcs<>"" Then Rs("CuoS")=strjzcs
       	Rs.update
   	Rs.Close	
   	Call WriteSuccessMsg("流水号为 <strong>" & strlsh &  "</strong> 的问题分析添加成功!","")
End Function
%>
