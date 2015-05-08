<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<style type="text/css">
<!--
.STYLE1 {color: #EEEEEE}
-->
</style>
<!-- #include file="Head.asp" -->
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
        <!--#include file="Left_index.asp" -->
	</td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
        <td background="Images_v1/yx2_2.gif">&nbsp;</td>
        <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
      </tr>
    </table>
        <a href="Product.asp"></a>
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td   width="132" valign="top" ><img src="Images_v1/book.gif" width="132" height="249"></td>
            <td width="6"></td>
            <td valign="top"><table  width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td height="206" valign="top"><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                      <TR>
                        <TD><table width="100%" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr>
                              <td width="29%" height="18">　留言中心&gt;&gt; 
                                留言专区</td>
                              <td width="71%">&nbsp;</td>
                            </tr>
                            <tr>
                              <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
                            </tr>
                            <tr>
                              <td height="1" colspan="2"><table border="0" align="right" cellpadding="0" cellspacing="3">
                                  <tr>
                                    <td><a href="Feedback.asp"><img src="Images_v1/write.gif" width="81" height="22" border="0"></a></td>
                                    <td><a href="FeedbackView.asp"><img src="Images_v1/read.gif" alt="查看留言" width="79" height="25" border="0"></a></td>
                                  </tr>
                              </table></td>
                            </tr>
                          </table>
                            <%
set rs=server.createobject("adodb.recordset")
sql="select * from Feedback where Publish<>'1' order by id desc"
rs.open sql,conn,1,1

dim MaxPerPageFeedback
MaxPerPageFeedback=5

'取得页数,并判断留言输入的是否数字类型的数据，如不是将以第一页显示
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

'显示帖子的子程序
Sub list()%>
                            <%
If rs.eof and rs.bof then
  response.write "<br>&nbsp;&nbsp;&nbsp;&nbsp;没有留言"
  exit Sub
End if
%>
                            <%
i=0
do while not rs.eof
%>
                            <table width="100%"  border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#E3E8EC" >
                              <tr bgcolor="#DFDFDF">
                                <td width="17%" height="22" align="center" background="Images_v1/book_bg.gif" bgcolor="#E3E8EC" ><img src="Images_v1/book_tt.gif" width="10" height="22" hspace="5" align="absmiddle">主　题： </td>
                                <td colspan="5" background="Images_v1/book_bg.gif" bgcolor="#E3E8EC">&nbsp;
                                    <%if rs("username")="管理员"then%>
                                  [管理员公告]
                                  <%end if%>
                                  <%=rs("title")%></a></td>
                              </tr>
                              <tr bgcolor="#DFDFDF">
                                <td height="22" align="right" bgcolor="#F4F7F7" > 反馈内容： </td>
                                <td colspan="5" align="center" bgcolor="#F4F7F7" ><table width="97%"  border="0" cellpadding="0" cellspacing="0">
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
                              <tr bgcolor="#F9FBFB">
                                <td height="22" align="right" bgcolor="#F4F7F7" > 留言者： </td>
                                <td width="22%" bgcolor="#F4F7F7" ><%=rs("Receiver")%> </td>
                                <td width="12%" align="right" bgcolor="#F4F7F7" >留言时间：</td>
                                <td width="17%" bgcolor="#F4F7F7" ><%=rs("Time")%></td>
                                <td width="12%" align="right" bgcolor="#F4F7F7" >回复时间：</td>
                                <td width="20%" bgcolor="#F4F7F7" ><%=rs("Retime")%> </td>
                              </tr>
                              <tr bgcolor="#DFDFDF">
                                <td height="22" align="right" bgcolor="#F4F7F7" > 管理员回复： </td>
                                <td colspan="5" align="center" bgcolor="#F4F7F7" ><table width="97%"  border="0" cellpadding="0" cellspacing="0" >
                                    <tr>
                                      <td height="4"></td>
                                    </tr>
                                    <tr>
                                      <td height="40" ><%if rs("ReFeedback")=""then%>
                                        ["未回复"]
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
                            </table>
                          <table width="100%" height="8" border="0" cellpadding="0" cellspacing="0">
                              <tr>
                                <td></td>
                              </tr>
                            </table>
                          <%
i=i+1
if i >= MaxPerPageFeedback then exit do
rs.movenext
loop
%>
                            <table width="100%" border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td><div align="right">
                                    <%
Response.write "全部"
Response.write "共" & Cstr(Rs.RecordCount) & "条留言&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "第" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='FeedbackView.asp?&page="+cstr(1)+"'>&nbsp;首页&nbsp;</a>"
Response.write "<a href='FeedbackView.asp?page="+Cstr(currentpage-1)+"'>&nbsp;上一页&nbsp;</a>"
Else
Response.write "&nbsp;上一页&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='FeedbackView.asp?page="+Cstr(currentPage+1)+"'>&nbsp;下一页&nbsp;</a>"
Response.write "<a href='FeedbackView.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;尾页&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;下一页&nbsp;"
End if
Response.write "转到第"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='FeedbackView.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='FeedbackView.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>页"
%>
                                </div></td>
                              </tr>
                            </table>
                          <%end sub%>
                        </TD>
                      </TR>
                  </table></td>
                </tr>
            </table></td>
          </tr>
        </table></td>
  </tr>
</table>
<%
rs.close
set rs=nothing
%>    
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>