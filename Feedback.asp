<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<%
UserName=Session("UserName")
set rs = Server.CreateObject("ADODB.recordset")
sql="select * from User where UserName='"&UserName&"'"
rs.open sql,conn,1,1
%>
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
                                  <td><a href="Feedback.asp"><img src="Images_v1/write.gif" alt="签写留言" width="81" height="22" border="0"></a></td>
                                  <td><a href="FeedbackView.asp"><img src="Images_v1/read.gif" alt="查看留言" width="79" height="25" border="0"></a></td>
                                </tr>
                            </table></td>
                          </tr>
                        </table>
                          <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                              <td height="206" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> </tr>
                                </table>
                                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td><form method="post" action="FeedbackSave.asp">
                                          <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td>尊敬的客户：
                                                <p>　　如果您对我们的产品或服务有任何意见和建议请及时告诉我们，我们将尽快给您满意的答复。<br>
                                                  如果您注册一个会员号，以后每次留言时只要登录再不用重复填写你的联系信息了，并且通过会员管理可集中查看只属于您的留言。（非注册会员也可直接留言） <br>
                                                </p></td>
                                            </tr>
                                          </table>
                                        <table width="98%" border="0" cellpadding="0" cellspacing="0">
                                            <tr>
                                              <td width="100%"><div align="center">
                                                  <table width="100%" height="409"
border="0" cellpadding="2" cellspacing="1" bgcolor="#E3E6E3">
                                                    <tr>
                                                      <input type=hidden name=Username value=<%=Username%>>
                                                      <td height="25" align="right" bgcolor="#F9FBFB">主题： </td>
                                                      <td height="25" bgcolor="#F9FBFB"><input type="text" name="Title" size="45" maxlength="36" style="font-size: 14px" >
                                                        *</td>
                                                    </tr>
                                                    <tr>
                                                      <td height="25" align="right" bgcolor="#F9FBFB">内容：</td>
                                                      <td height="25" bgcolor="#F9FBFB"><textarea rows="10" name="Content" cols="45" style="font-size: 14px" ></textarea></td>
                                                    </tr>
                                                    <tr>
                                                      <td height="25" align="right" bgcolor="#F9FBFB">悄悄话：</td>
                                                      <td width="83%" height="11" bgcolor="#F9FBFB"><% if session("username")<>"" then  %>
                                                          <input type="radio" name="Publish" value="1">
                                                        是
                                                        <input name="Publish" type="radio" value="0" checked>
                                                        否
                                                        <% else %>
                                                        <input name="Publish" type="hidden" value="0" checked>
                                                        <%	  response.write "<font color='#ff6600'>非注册会员不可设</font>"
					end if%>
                                                        (只有管理员和自己能看)</td>
                                                    </tr>
                                                    <tr>
                                                      <td width="17%" height="0" valign="top" bgcolor="#F9FBFB"></td>
                                                      <td width="83%" height="0" valign="top" bgcolor="#F9FBFB"><input type="submit" value="提交留言"
name="cmdOk">
                                                          <input type="reset" value="重写" name="cmdReset">
                                                      </td>
                                                    </tr>
                                                  </table>
                                              </div></td>
                                            </tr>
                                          </table>
                                      </form></td>
                                    </tr>
                                </table></td>
                            </tr>
                        </table></TD>
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
