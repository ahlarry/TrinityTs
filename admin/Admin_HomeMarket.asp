<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%sale=server.htmlencode(Trim(Request("content"))) %>
<%if Request.QueryString("mark")="southidc" then
set rs=server.createobject("adodb.recordset")
sql="select * from Market2"
rs.open sql,conn,3,3
rs("HomeMarket")=sale
rs.update
rs.close
response.redirect "Admin_HomeMarket.asp"
end if
sql="select * from Market2"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <strong><br>
      </strong> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td height="25" class="back_southidc"> 
            <div align="center"><strong>国 内 市 场 设 置</strong></div></td>
        </tr>
        <tr> 
          <form method="POST" action="Admin_HomeMarket.asp?mark=southidc">
            <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                <tr> 
                  <td height="300" bgcolor="#ECF5FF">  <textarea name="Content" style="display:none"><%=rs_home("HomeMarket")%></textarea> 
                      <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="450"></iframe> </td>
                </tr>
                <tr bgcolor="#ECF5FF"> 
                  <td height="30"> <div align="center"> 
                      <input type="submit" value=" 修 改 "
name="cmdok">
                      &nbsp; 
                      <input type="reset" value=" 重 写 "
name="cmdcancel">
                    </div></td>
                </tr>
              </table></td>
          </form>
        </tr>
      </table>
      <strong> </strong></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->