<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<%
	Dim strlsh, stryscxs, strytsxs, strscxs, strtsxs
	strlsh=Request("ilsh")		'��ˮ��
	stryscxs=Request("yscxs") 	'ԭ����ϵ��
	strytsxs=Request("ytsxs") 	'ԭ����ϵ��
	strscxs=Request("scxs")		'������ϵ��
	strtsxs=Request("tsxs")		'������ϵ��

	If strscxs<>"" Then
		strSql="select * from [ins_tsfz] where lsh='"&strlsh&"'and js='����'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		Do While Not Rs.eof
			If Rs("fz")<>Round(Rs("fz")/stryscxs*strscxs,1) Then
				Rs("bz")=Rs("fz")&"("&stryscxs&"��"&strscxs&"),"&now()
				Rs("fz")=Round(Rs("fz")/stryscxs*strscxs,1)
			End If
			Rs.movenext
		loop
	End If
	If strtsxs<>"" Then
		strSql="select * from [ins_tsfz] where lsh='"&strlsh&"'and js='����'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,3
		Do While Not Rs.eof
			If Rs("fz")<>Round(Rs("fz")/strytsxs*strtsxs,1) Then
				Rs("bz")=Rs("fz")&"("&strytsxs&"��"&strtsxs&"),"&now()
				Rs("fz")=Round(Rs("fz")/strytsxs*strtsxs,1)
			End If
			Rs.movenext
		loop
	End If

	Call WriteSuccessMsg("��ˮ�� " & strlsh &  " ���������ֵϵ���������!","")
	Rs.Close
%>
