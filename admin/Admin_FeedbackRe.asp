<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!--#include file="inc/articlechar.inc"-->

<%if Request.QueryString("mark")="southidc" then
id=request("id")

if request.form("html")="on" then
ReFeedback=request.form("ReFeedback")
else
ReFeedback=htmlencode2(request.form("ReFeedback"))
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from Feedback where id="&id
rs.open sql,conn,1,3

rs("ReFeedback")=ReFeedback
rs("ReTime")=date()
rs.update
rs.close
response.redirect "Admin_Feedback.asp"
end if

id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From Feedback where id="&id, conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <form method="post" action="Admin_FeedbackRe.asp?mark=southidc">
      <input type=hidden name=id value=<%=id%>>
      <td align="center" valign="top"> <br> <br> 
        <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
          <tr> 
            <td height="25" class="back_southidc"> 
              <div align="center"><strong>�ظ����� <br>
                </strong></div></td>
          </tr>
          <tr> 
            <td><div align="center"> 
                <table width="100%" border="0" cellpadding="0" cellspacing="1" bgcolor="#000000">
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25" colspan="2"> 
                      <div align="center"></div>
                      &nbsp;&nbsp;����׼���ظ�<b><%=rs("name")%></b>�����ԣ����������ǣ�</td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td width="18%" height="22"> 
                      <div align="center">�� ��</div></td>
                    <td width="82%" bgcolor="#ECF5FF">&nbsp;&nbsp;<%=rs("title")%></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="25"> 
                      <div align="center">�� ��</div></td>
                    <td height="25">&nbsp;&nbsp;<%=rs("content")%></td>
                  </tr>
                </table>
              </div></td>
          </tr>
        </table>
        <br> <br> 
        <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
          <tr> 
            <td class="back_southidc" height="25"> 
              <div align="center"><strong>����ظ�����</strong></div></td>
          </tr>
          <tr> 
            <td bgcolor="#ECF5FF">
<div align="center"> 
                <textarea name="ReFeedback" cols="70" rows="10"><%=rs("ReFeedback")%></textarea>
              </div></td>
          </tr>
          <tr> 
            <td bgcolor="#ECF5FF">
<div align="center">�Ƿ�֧��Html 
                <input type="checkbox" name="html" value="on">
                <input type="submit" value="ȷ��">
                <input type="reset" value="����">
              </div></td>
          </tr>
        </table></td>
    </form>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->