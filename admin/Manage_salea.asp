<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%salea=server.htmlencode(Trim(Request("content"))) %>
<%if Request.QueryString("no")="eshop" then
set rs=server.createobject("adodb.recordset")
sql="select * from main"
rs.open sql,conn,3,3
rs("salea")=salea
rs.update
rs.close
response.redirect "Manage_salea.asp"
end if
sql="select * from main"
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
            <div align="center"><strong>�� �� �� �� �� ��</strong></div></td>
        </tr>
        <tr> 
          <form method="POST" action="Manage_salea.asp?no=eshop">
            <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                <tr> 
                  <td width="19%" height="300" bgcolor="#ECF5FF"> <div align="right">����ָ�ϣ�<br>
                      ֧��UBB���� </div></td>
                  <td width="81%" bgcolor="#ECF5FF"> <script src=../Inc/Eshopcode.js></script> 
                    <!--#include file=../Inc/Ubb.inc -->
                    <textarea name="content" cols="58" rows="15">
<%
salea=rs_home("salea")
response.write salea%>
</textarea> </td>
                </tr>
                <tr bgcolor="#ECF5FF"> 
                  <td height="30" colspan="2"> <div align="center"> 
                      <input type="submit" value=" �� �� "
name="cmdok">
                      &nbsp; 
                      <input type="reset" value=" �� д "
name="cmdcancel">
                    </div></td>
                </tr>
                <tr> 
                  <td colspan="2">ͼƬ�ϴ�</td>
                </tr>
                <tr> 
                  <td colspan="2"><div align="center"> 
                      <iframe name="ad" frameborder=0 width=100% height=50 scrolling=no src=../upload_Other.asp></iframe>
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