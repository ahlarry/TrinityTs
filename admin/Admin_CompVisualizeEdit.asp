<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<%if Request.QueryString("mark")="southidc" then

id=request("id")
Explain=server.htmlencode(Trim(request("Explain")))
CompVisualize=server.htmlencode(Trim(Request("CompVisualize")))

If explain="" Then
response.write "SORRY <br>"
response.write "�������Ʋ���Ϊ��!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if
If CompVisualize="" Then
response.write "SORRY <br>"
response.write "����ͼƬ����Ϊ��!<a href=""javascript:history.go(-1)"">��������</a>"
response.end
end if

Set rs = Server.CreateObject("ADODB.Recordset")
sql="select * from CompVisualize where id="&id
rs.open sql,conn,1,3

rs("Explain")=Explain
rs("CompVisualize")=CompVisualize
rs("Time")=Date()
rs.update
rs.close
response.redirect "Admin_CompVisualize.asp"
end if
%>
<%
id=request.querystring("id")
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From CompVisualize where id="&id, conn,1,1
%>
<!-- #include file="Inc/Head.asp" -->
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"> <br> <table width="560" border="0" cellpadding="2" cellspacing="1" class="table_southidc">
        <tr> 
          <td class="back_southidc" height="25"> <div align="center"><strong>�޸�����<br>
              </strong></div></td>
        </tr>
        <tr> 
          <form method="post" name="myform" action="Admin_CompVisualizeEdit.asp?mark=southidc">
            <input type=hidden name=id value=<%=id%>>
            <td><div align="center"> 
                <table width="100%" border="0" cellspacing="1" cellpadding="0">
                  <tr bgcolor="#ECF5FF"> 
                    <td width="24%" height="25"> 
                      <div align="center">����˵��</div></td>
                    <td width="76%"> 
                      <textarea name="Explain" cols="40" rows="8"><%=rs("Explain")%></textarea></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="22"> 
                      <div align="center">����ͼƬ</div></td>
                    <td> 
                      <input name="CompVisualize" type="text" id="img" value="<%=rs("CompVisualize")%>" size="40" maxlength="50"> 
                      <font color="#FF0000"> * ͼƬ��ַ</font></td>
                  </tr>
                  <tr bgcolor="#ECF5FF"> 
                    <td height="12" colspan="2"> 
                      <div align="center"> 
                        <input name="submit" type="submit" value="ȷ��">
                        &nbsp; 
                        <input name="reset" type="reset" value="����">
                      </div></td>
                  </tr>
                  <tr bgcolor="#ECF5FF">
                    <td height="12" colspan="2">������ҳ��
                      <input name="ishome" type="checkbox" id="ishome" value="yes" <% if rs("ishome")=True then Response.Write("checked") %>>
(��ѡ���ͼ������ҳ��ϵͳ���������ϴ����޸ĵ�ͼƬ����ҳ)</td>
                  </tr>
                  <tr> 
                    <td colspan="2">ͼƬ�ϴ�</td>
                  </tr>
                  <tr> 
                    <td colspan="2"><div align="left"> 
                         <iframe style="top:2px" ID="UploadFiles" src="../upload_Photo.asp?PhotoUrlID=3" frameborder=0 scrolling=no width="300" height="25"></iframe>
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