<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<SCRIPT language=javascript>
function unselectall()
{
    if(document.del.chkAll.checked){
	document.del.chkAll.checked = document.del.chkAll.checked&0;
    } 	
}

function CheckAll(form)
  {
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll")
       e.checked = form.chkAll.checked;
    }
  }
function ConfirmDel()
{
   if(confirm("ȷ��Ҫɾ��ѡ�е���Ŀ��һ��ɾ�������ָܻ���"))
     return true;
   else
     return false;	 
}

</SCRIPT>

<%if Request.QueryString("mark")="southidc" then
dim SQL, Rs, ID,CurrentPage
CurrentPage = request("Page")
ID=request("ID")

set rs=server.createobject("adodb.recordset")
sqltext="delete from CompVisualize where Id="& ID
rs.open sqltext,conn,1,3
set rs=nothing
conn.close
response.redirect "Admin_CompVisualize.asp"
end if
%>
<%
set rs=server.createobject("adodb.recordset")
sqltext="select * from CompVisualize order by id desc"
rs.open sqltext,conn,1,1

dim MaxPerPage
MaxPerPage=10

dim text,checkpage
text="0123456789"
Rs.PageSize=MaxPerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
If NOT IsEmpty(request("page")) Then
CurrentPage=Cint(request("page"))
If CurrentPage < 1 Then CurrentPage = 1
If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount
Else
CurrentPage= 1
End If
If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
CurrentPage=1
End if

call list

'��ʾ���ӵ��ӳ���
Sub list()%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>��ҵ������� <br>
              </strong></div></td>
        </tr>
        <tr> 
          <td><div align="center"> 
              <table width="100%" border="0" cellspacing="2" cellpadding="3">
                <tr bgcolor="#A4B6D7"> 
                  <td width="8%" height="25"> 
                    <div align="center">ID</div></td>
                  <td width="25%"> 
                    <div align="center">��������</div></td>
                  <td width="47%"> 
                    <div align="center">����ͼƬ</div></td>
                  <td width="10%"> 
                    <div align="center">����</div></td>
                  <td width="10%"> 
                    <div align="center">����</div></td>
                </tr>
                <%
if not rs.eof then
i=0
do while not rs.eof
%>
                <tr bgcolor="#ECF5FF"> 
                  <td height="22"> 
                    <div align="center"><%=rs("id")%></div></td>
                  <td>&nbsp;&nbsp;<%=rs("Explain")%></td>
                  <td> 
                    <div align="center"><img name="" src="../<%=rs("CompVisualize")%>" width="100" height="120" alt=""></div></td>
                  <td> 
                    <div align="center"><a href="Admin_CompVisualizeEdit.asp?id=<%=rs("id")%>">�޸�</a></div></td>
                  <td> 
                    <div align="center"><a href="Admin_CompVisualize.asp?id=<%=rs("id")%>&mark=southidc" onclick="return ConfirmDel();">ɾ��</a></div></td>
                </tr>
                <% 
i=i+1 
if i >= MaxPerpage then exit do 
rs.movenext 
loop 
end if 
%>
                <tr bgcolor="#A4B6D7"> 
                  <td height="25" colspan="5"> <div align="right">
                      <%
Response.write "ȫ��"
Response.write "��" & Cstr(Rs.RecordCount) & "������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "��" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='Admin_CompVisualize.asp?&page="+cstr(1)+"'>&nbsp;��ҳ&nbsp;</a>"
Response.write "<a href='Admin_CompVisualize.asp?page="+Cstr(currentpage-1)+"'>&nbsp;��һҳ&nbsp;</a>"
Else
Response.write "&nbsp;��һҳ&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='Admin_CompVisualize.asp?page="+Cstr(currentPage+1)+"'>&nbsp;��һҳ&nbsp;</a>"
Response.write "<a href='Admin_CompVisualize.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;βҳ&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;��һҳ&nbsp;"
End if
Response.write "ת����"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='Admin_CompVisualize.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='Admin_CompVisualize.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>ҳ"
%>
                    </div></td>
                </tr>
              </table>
<%
End sub
rs.close
%>
            </div></td>
        </tr>
      </table>
      <br> <br> </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->