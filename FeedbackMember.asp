<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!-- #include file="Head.asp" -->
<%
UserName=trim(request("UserName"))
if session("UserName")="" then
	response.redirect "UserServer.asp"
end if
%>
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="left_usres.asp" -->
        <!--#include file="left_index.asp" -->
    </td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
        <td background="Images_v1/yx2_2.gif">&nbsp;</td>
        <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
      </tr>
    </table>
        <a href="Product.asp"></a>
        <table width="96%" border="0" cellspacing="2" cellpadding="0">
          <TR>
            <TD><table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
                <tr>
                  <td width="34%">&nbsp;&nbsp;<a href="index.asp">��ҳ</a> &gt;&gt; �ͻ����� 
                    &gt;&gt; ��Ա����ר��</td>
                  <td width="66%">&nbsp;</td>
                </tr>
                <tr>
                  <td colspan="2" bgcolor="#999999"></td>
                </tr>
              </table>
                <%
UserName=Session("UserName")
set rs=server.createobject("adodb.recordset")
sql="select * from Feedback where UserName='"&UserName&"' order by id desc"
rs.open sql,conn,1,1

dim MaxPerPageFeedback
MaxPerPageFeedback=5

'ȡ��ҳ��,���ж�����������Ƿ��������͵����ݣ��粻�ǽ��Ե�һҳ��ʾ
dim text,checkpage
text="0123456789"
Rs.PageSize=MaxPerPageFeedback
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
                <%
If rs.eof and rs.bof then
  response.end
End if
%>
                <br>
                <%
i=0
do while not rs.eof
%>
                <table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E6EAED" >
                  <tr bgcolor="#DFDFDF">
                    <td width="17%" height="22" align="right" background="Images_v1/yx_bg.gif" bgcolor="#DFDFDF" > �����⣺ </td>
                    <td colspan="5" background="Images_v1/yx_bg.gif" bgcolor="#DFDFDF">&nbsp;
                        <%if rs("username")="����Ա"then%>
                      [����Ա����]
                      <%end if%>
                      <%=rs("title")%></a></td>
                  </tr>
                  <tr bgcolor="#DFDFDF">
                    <td height="22" align="right" bgcolor="#F8F9F8" > �������ݣ� </td>
                    <td colspan="5" align="center" bgcolor="#F8F9F8" ><table width="97%"  border="0" cellpadding="0" cellspacing="0">
                        <tr>
                          <td height="4">&nbsp;</td>
                        </tr>
                        <tr>
                          <td height="40" ><%=rs("Content")%> </td>
                        </tr>
                        <tr>
                          <td height="4"></td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr bgcolor="#EFEFEF">
                    <td height="22" align="right" bgcolor="#F8F9F8" > �����ߣ� </td>
                    <td width="22%" bgcolor="#F8F9F8" ><%=rs("Receiver")%> </td>
                    <td width="12%" align="right" bgcolor="#F8F9F8" >����ʱ�䣺</td>
                    <td width="17%" bgcolor="#F8F9F8" ><%=rs("Time")%></td>
                    <td width="12%" align="right" bgcolor="#F8F9F8" >�ظ�ʱ�䣺</td>
                    <td width="20%" bgcolor="#F8F9F8" ><%=rs("Retime")%> </td>
                  </tr>
                  <tr bgcolor="#DFDFDF">
                    <td height="22" align="right" bgcolor="#EFEFEF" > ����Ա�ظ��� </td>
                    <td colspan="5" align="center" bgcolor="#EFEFEF" ><table width="97%"  border="0" cellpadding="0" cellspacing="0" >
                        <tr>
                          <td height="4"></td>
                        </tr>
                        <tr>
                          <td height="40" ><%if rs("ReFeedback")=""then%>
                            ["δ�ظ�"]
                            <%else%>
                            <%=rs("ReFeedback")%>
                            <%end if%>
                          </td>
                        </tr>
                        <tr>
                          <td height="4"></td>
                        </tr>
                    </table></td>
                  </tr>
                  <tr valign="top" bgcolor="#F7F7F7">
                    <td height="12" colspan="6" >&nbsp;</td>
                  </tr>
                </table>
              <%
i=i+1
if i >= MaxPerPageFeedback then exit do
rs.movenext
loop
%>
                <table width="100%" border="0" cellspacing="3" cellpadding="0">
                  <tr>
                    <td><div align="right">
                        <%
Response.write "ȫ��"
Response.write "��" & Cstr(Rs.RecordCount) & "������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "��" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='FeedbackMember.asp?&page="+cstr(1)+"'>&nbsp;��ҳ&nbsp;</a>"
Response.write "<a href='FeedbackMember.asp?page="+Cstr(currentpage-1)+"'>&nbsp;��һҳ&nbsp;</a>"
Else
Response.write "&nbsp;��һҳ&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='FeedbackMember.asp?page="+Cstr(currentPage+1)+"'>&nbsp;��һҳ&nbsp;</a>"
Response.write "<a href='FeedbackMember.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;βҳ&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;��һҳ&nbsp;"
End if
Response.write "ת����"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='FeedbackMember.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='FeedbackMember.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>ҳ"
%>
                    </div></td>
                  </tr>
              </table>
              <%end sub%></TD>
          </TR>
        </table></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>    
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>