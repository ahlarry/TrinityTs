<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
option explicit
response.buffer=true	
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="function.asp"-->
<html>
<head>
<meta http-equiv=Content-Type content="text/html; charset=gb2312">
<style type="text/css">body {font-size:	9pt}</style>
</head>
<BODY bgcolor="#FFFFFF" MONOSPACE>
<%
dim Action,FoundErr,ErrMsg
dim ArticleID,sql,rs,Tb
Action=trim(request("Action"))
ArticleID=trim(request("ArticleID"))
Tb=request("Tb")
if Action="Modify" then
	if ArticleId="" then
		response.write "��ָ��Ҫ�޸ĵ�����ID"
	else
		if Tb="products" then
		 sql="select Content from "&Tb&" where ArticleID=" & ArticleID & ""
		end if
		if Tb="market" then
		 sql="select Content from "&Tb&" where id=" & ArticleID & ""
		end if
		Set rs= Server.CreateObject("ADODB.Recordset")
		rs.open sql,conn,1,1
		if rs.bof and rs.eof then
			response.write "�Ҳ�����Ʒ"
		else
			response.write "<p>" & rs("Content") & "</p>"
		end if
		rs.close
		set rs=nothing
	end if
end if
%>
</body>
</html>
<%
call CloseConn()
%>