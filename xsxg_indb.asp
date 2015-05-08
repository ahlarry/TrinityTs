<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim strlsh, stryscxs, strytsxs, strscxs, strtsxs
	strlsh=Request("ilsh")		'流水号
	stryscxs=Request("yscxs") 	'原生产系数
	strytsxs=Request("ytsxs") 	'原调试系数
	strscxs=Request("scxs")		'生产线系数
	strtsxs=Request("tsxs")		'调试线系数

	If strscxs<>"" Then
		strSql="select * from [ins_tsfz] where lsh='"&strlsh&"'and js='生产'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		Do While Not Rs.eof
			If Rs("fz")<>Round(Rs("fz")/stryscxs*strscxs,1) Then
				Rs("bz")=Rs("fz")&"("&stryscxs&"→"&strscxs&"),"&now()
				Rs("fz")=Round(Rs("fz")/stryscxs*strscxs,1)
			End If
			Rs.movenext
		loop
	End If
	If strtsxs<>"" Then
		strSql="select * from [ins_tsfz] where lsh='"&strlsh&"'and js='调试'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		Do While Not Rs.eof
			If Rs("fz")<>Round(Rs("fz")/strytsxs*strtsxs,1) Then
				Rs("bz")=Rs("fz")&"("&strytsxs&"→"&strtsxs&"),"&now()
				Rs("fz")=Round(Rs("fz")/strytsxs*strtsxs,1)
			End If
			Rs.movenext
		loop
	End If

	Call WriteSuccessMsg("流水号 " & strlsh &  " 厂内任务分值系数调整完成!","")
	Rs.Close
%>
