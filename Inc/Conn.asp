<%
Response.Expires=0
Response.Buffer=True
Session.Timeout=600
Response.Expires = -1
Response.CacheControl="no-chache"
Response.ExpiresAbsolute= Now() - 100

dim conn,db
dim connstr
db="Databases/#2007debug.mdb" '���ݿ��ļ�λ��
on error resume next
connstr="DBQ="+server.mappath(""&db&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if
sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>