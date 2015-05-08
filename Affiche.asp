<!--#include file="Inc/conn.asp"-->
<!--#include file="inc/ubbcode.asp"-->
<%
dim ID
dim sqlAnnounce
dim rsAnnounce
dim AnnounceNum
dim ChannelID
ID=Trim(request("ID"))
sqlAnnounce="select * from affiche where ID=" & Clng(ID)
sqlAnnounce=sqlAnnounce & " order by ID Desc"
Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
rsAnnounce.open sqlAnnounce,conn,1,1
%>
<html>
<head>
<title>本站公告</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
a:active { text-decoration: none; color: #0000FF}
a:hover { text-decoration: none; color: #FF0000}
a:link { text-decoration: none; color: #0000FF}
a:visited { text-decoration: none; color: #990000}
BODY { text-decoration: none; font-size: 12px}
TABLE { text-decoration: none; font-size: 12px}
</style>
</head>
<BODY text=#000000 bgColor=#ffffff leftMargin=0 topMargin=0>
<TABLE width=400 border=0 cellPadding=0 cellSpacing=0 background="Images_v1/aff2.gif">
  <TBODY>
    <TR > 
      <TD width="23" rowspan="3" valign="top" ><img src="Images_v1/aff1.gif" width="23" height="300"></TD>
      <TD width="358" align="center"><img src="Images_v1/aff3_till.gif" alt="公告" width="86" height="25" vspace="5"></TD>
      <TD width="19" rowspan="3" valign="top" ><img src="Images_v1/aff3.gif" width="19" height="300"></TD>
    </TR>
    
    <TR> 
      <TD vAlign=top> <TABLE width="100%" height="232" border=0 cellPadding=3 cellSpacing=5>
          <TBODY>
            <TR> 
              <TD valign="top"><%
if rsAnnounce.bof and rsAnnounce.eof then 
	response.write "<p>&nbsp;&nbsp;没有通告或找不到指定的通告</p>" 
else 
	AnnounceNum=rsAnnounce.recordcount
	dim i
	do while not rsAnnounce.eof
		response.Write "<p align='center'>" & rsAnnounce("title") & "</p><p align='left'>" & ubbcode(dvHTMLEncode(rsAnnounce("Content"))) & "</p><p align='right'>" & "&nbsp;&nbsp;<br>" & FormatDateTime(rsAnnounce("Time"),1) & "</p>"
		rsAnnounce.movenext
		i=i+1
		if i<AnnounceNum then response.write "<hr>"
	loop	
end if  
%></TD>
            </TR>
          </TBODY>
      </TABLE></TD>
    </TR>
    
    <TR> 
      <TD><div align="right"><a href="javascript:window.close();"><img src="Images_v1/aff3_close.gif" alt="关闭窗口" width="67" height="21" border="0"></a></div></TD>
    </TR>
  </TBODY>
</TABLE>
<%
rsAnnounce.close
set rsAnnounce=nothing
call CloseConn()
%>
<map name="Annouce_r6_c1Map">
  <area shape="rect" coords="170,13,230,27">
</map>
