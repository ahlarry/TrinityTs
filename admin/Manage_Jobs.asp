<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%Job=server.htmlencode(Trim(Request("content"))) %>
<%if Request.QueryString("mark")="southidc" then
set rs=server.createobject("adodb.recordset")
sql="select * from main"
rs.open sql,conn,3,3
rs("Job")=Job
rs.update
rs.close
response.redirect "Manage_Jobs.asp"
end if
sql="select * from main"
Set rs_home= Server.CreateObject("ADODB.Recordset")
rs_home.open sql,conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="862" align="center" valign="top"> <strong><br>
      </strong> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td height="25" class="back_southidc"> 
            <div align="center"><strong>人 才 策 略</strong> <strong>设 置</strong></div></td>
        </tr>
        <tr> 
          <form method="POST" action="Manage_Jobs.asp?mark=southidc">
            <td><table width="100%" border="0" cellspacing="1" cellpadding="0">
                <tr> 
                  <td width="19%" bgcolor="#ECF5FF"> 
                    <div align="right">付费指南：<br>
                      支持UBB代码 </div></td>
                  <td width="81%" bgcolor="#ECF5FF"> 
                    <script src=../Inc/Eshopcode.js></script> 
                    <!--#include file=../Inc/Ubb.inc -->
                    <textarea name="content" cols="58" rows="15"><%if rs_home("html")=false then
Job=replace(rs_home("Job"),"<br>",chr(13))
Job=replace(Job,"&nbsp;"," ")
else
Job=rs_home("Job")
end if
response.write Job%></textarea>
                  </td>
                </tr>
                <tr bgcolor="#ECF5FF"> 
                  <td height="30" colspan="2"> 
                    <div align="center"> 
                      <input type="submit" value=" 修 改 "
name="cmdok">
                      &nbsp; 
                      <input type="reset" value=" 重 写 "
name="cmdcancel">
                    </div></td>
                </tr>
                <tr> 
                  <td colspan="2">图片上传</td>
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