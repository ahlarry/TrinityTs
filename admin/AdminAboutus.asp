<%@language=VBScript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="Inc/Function.asp"-->
<%
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function%>
<%
dim strFileName
const MaxPerPage=20
dim totalPut,CurrentPage,TotalPages
dim rs, sql
strFileName="AdminAboutus.asp"

if request("page")<>"" then
    currentPage=cint(request("page"))
else
	currentPage=1
end if

Set rs=Server.CreateObject("Adodb.RecordSet")
sql="select * from Aboutus order by Aboutusorder"
rs.Open sql,conn,1,1
%>
<script language=javascript>
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ���˼�¼��"))
     return true;
   else
     return false;
}
</script>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top">
      <br>
      <%
  	if rs.eof and rs.bof then
		response.write "Ŀǰ���� 0 ����¼"
	else
    	totalPut=rs.recordcount
		if currentpage<1 then
       		currentpage=1
    	end if
    	if (currentpage-1)*MaxPerPage>totalput then
	   		if (totalPut mod MaxPerPage)=0 then
	     		currentpage= totalPut \ MaxPerPage
		  	else
		      	currentpage= totalPut \ MaxPerPage + 1
	   		end if

    	end if
	    if currentPage=1 then        
        	showContent
        	showpage strFileName,totalput,MaxPerPage,true,true,"����¼"
   	 	else
   	     	if (currentPage-1)*MaxPerPage<totalPut then
         	   	rs.move  (currentPage-1)*MaxPerPage
         		dim bookmark
           		bookmark=rs.bookmark        		
            	showContent
            	showpage strFileName,totalput,MaxPerPage,true,true,"����¼"
        	else
	        	currentPage=1        
           		showContent
           		showpage strFileName,totalput,MaxPerPage,true,true,"����¼"
	    	end if
		end if
	end if

sub showContent
   	dim i
    i=0
%> 
      <table width="556" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr>
          <td width="550" height="25" class="back_southidc"> 
            <div align="center"><strong>��Ŀ����</strong></div></td>
        </tr>
        <tr> 
          <td height="66"> 
            <div align="center">
              <table width="100%" height="62" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#000000" class="border">
                <tr bgcolor="#ECF5FF" class="title">
                  <td width="161" align="center"><strong>�����</strong></td>
                  <td width="249" height="29" align="center"><strong> ��Ŀ����</strong></td>
                  <td width="136" align="center"><strong> ����</strong></td>
                </tr>
                <%do while not rs.EOF %>
                <tr class="tdbg">
                  <td align="center" bgcolor="#ECF5FF"><%=rs("Aboutusorder")%></td>
                  <td height="28" align="center" bgcolor="#ECF5FF"><%=rs("Title")%></td>
                  <td align="center" bgcolor="#ECF5FF"><a href=../Aboutus.asp?Title=<%=rs("Title")%> target="_blank">�鿴</a>&nbsp; 
                    &nbsp;<a href="AdminAboutusModify.asp?ID=<%=rs("ID")%>">�޸�</a>&nbsp; 
                    &nbsp;<a href="AdminAboutusDel.asp?ID=<%=rs("ID")%>" onClick="return ConfirmDel();">ɾ��</a></td>
                </tr>
                <%
	i=i+1
	if i>=MaxPerPage then exit do
	rs.movenext
loop
%>
              </table>
            </div></td>
        </tr>
      </table>
      <%
end sub 
%> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
rs.Close
set rs=Nothing
call CloseConn()  
%>