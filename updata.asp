<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/function.asp"-->
<title>���ݱ��</title>
  <%
	n=0	
	Dim strtsy,strlsh
	strtsy=1.1 : strlsh="5289"
	strSql="select * from [ins_tsfz] where lsh='"&strlsh&"' and zrr<>'װ��'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,conn,1,3
	do while not Rs.eof
		Response.Write("<br><font color='#ffff00'>"&Rs("zrr")&"</font><font color='#ff0000'>ԭ��ֵ="&Rs("fz")&"</font>")
  		Rs("fz")=Round(Rs("fz")*strtsy,1)
  		Rs("sj")=dateadd("m",-1,Rs("sj"))
  		Rs.update
  		Response.Write("<font color='#ff0000'>,�ַ�ֵ="&Rs("fz")&"</font>")
		rs.MoveNext
	loop
	rs.close
	Response.Write("<br><br><font color='#ff0000'>��ˮ��="&strlsh&"</font>")
%>