<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%if Request.QueryString("mark")="southidc" then
 OverseasMarket=Request("content")
 EnOverseasMarket=Request("Encontent")
 set rs=server.createobject("adodb.recordset")
 sql="select * from Market2"
 rs.open sql,conn,1,3
 rs("OverseasMarket")=OverseasMarket
 rs("EnOverseasMarket")=EnOverseasMarket
 rs.update
 rs.close
 response.redirect "Admin_OverseasMarket.asp"
end if
sql="select * from Market2"
Set rs_Market2= Server.CreateObject("ADODB.Recordset")
rs_Market2.open sql,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <strong><br>
      </strong> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td height="25" class="back_southidc"> 
            <div align="center"><strong>国 际 市 场 设 置</strong></div></td>
        </tr>
        <tr> 
          <form method="POST" action="Admin_OverseasMarket.asp?mark=southidc">
            <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                <tr> 
                  <td height="20" bgcolor="#ECF5FF"><div align="center">中文</div></td>
                </tr>
                <tr> 
                  <td height="300" bgcolor="#ECF5FF"> <input type="hidden" name="content" value="<%=Server.HTMLEncode(rs_Market2("OverseasMarket"))%>"> 
                    <iframe ID="eWebEditor1" src="Southidceditor/ewebeditor.asp?id=content&style=southidc" frameborder="0" scrolling="no" width="550" HEIGHT="450"></iframe></td>
                </tr>
                <tr bgcolor="#ECF5FF">
                  <td height="20"><div align="center"></div></td>
                </tr>
                <tr bgcolor="#ECF5FF"> 
                  <td>&nbsp;</td>
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
<%
rs_Market2.close
set rs_Market2=nothing
%>
<!-- #include file="Inc/Foot.asp" -->