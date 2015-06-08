<%
Response.Expires=0
Response.Buffer=True
Session.Timeout=600
Response.Expires = -1
Response.CacheControl="no-chache"
Response.ExpiresAbsolute= Now() - 100

dim conn,db,conn2,db2
dim connstr,connstr2
db="Databases/#2007debug.mdb" '数据库文件位置
on error resume next
connstr="DBQ="+server.mappath(""&db&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn.open connstr
end if

db2="Databases/#2012jsb.mdb" '数据库文件位置
on error resume next
connstr2="DBQ="+server.mappath(""&db2&"")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn2=server.createobject("ADODB.CONNECTION")
if err then
err.clear
else
conn2.open connstr2
end if

sub CloseConn()
	conn.close
	set conn=nothing
	conn2.close
	set conn2=nothing
end sub

%>