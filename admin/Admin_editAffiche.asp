<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%if Request.QueryString("mark")="southidc" then

id=request("id")
Title=request("Title")
Content=Request("Content")

If Title="" Then
response.write "SORRY <br>"
response.write "�����빫�����"
response.end
end if

If Content="" Then
response.write "SORRY <br>"
response.write "�������ݲ���Ϊ��"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Affiche where id="&id
rs.open sql,conn,1,3

rs("Title")=Title
rs("Content")=Content
rs.update
rs.close
response.redirect "Admin_AddAffiche.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Affiche where id="&id, conn,3,3
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <strong><br>
      </strong> <br> <br> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td height="25" class="back_southidc"> 
            <div align="center"><strong>�޸���վ����<br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" action="Admin_editAffiche.asp?mark=southidc">
            <input type=hidden name=id value=<%=id%>>
            <td><div align="center"> 
                <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#ECF5FF">
                  <tr> 
                    <td width="27%" height="25" bgcolor="#A4B6D7"> <div align="center">�������:</div></td>
                    <td width="73%"> <input name="Title" type="text" id="Title" value="<%=rs("Title")%>" size="45" maxlength="100"> 
                    </td>
                  </tr>
                  <tr> 
                    <td height="22" bgcolor="#A4B6D7"> <div align="center">��������:</div></td>
                    <td><textarea rows="8" name="Content" cols="45"><%=rs("Content")%></textarea> 
                    </td>
                  </tr>
                  <tr> 
                    <td height="25" colspan="2" bgcolor="#E3E3E3"> <div align="center"> 
                        <input name="submit" type="submit" value="ȷ��">
                        &nbsp; 
                        <input name="reset" type="reset" value="����">
                      </div></td>
                  </tr>
                </table>
              </div></td>
          </form>
        </tr>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->