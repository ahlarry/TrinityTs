<!--#include file="Inc/conn.asp" -->
<!--#include file="Inc/Function.asp" -->
<!--#include file="Inc/Config.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!-- #include file="Head.asp" -->
<!--#include file="inc/md5.asp"-->
<%
Action=Trim(request("Action"))
if Action="Login" then
 dim sql
 dim rs
 dim username
 dim password
 username=replace(trim(request("username")),"'","")
 password=replace(trim(Request("password")),"'","")
 password=md5(password)
 set rs=server.createobject("adodb.recordset")
 sql="select * from [User] where LockUser=False and username='" & username & "' and password='" & password &"'"
 rs.open sql,conn,1,3
 if not(rs.bof and rs.eof) then
	if password=rs("password") then
	    rs("LoginIP")=Request.ServerVariables("REMOTE_ADDR")
		rs("LastLoginTime")=now()
		rs("logins")=rs("logins")+1
		rs.update
		session("UserName")=rs("username")
		session("GZZ")=rs("gzz")		
		Response.Redirect "UserServer.asp"
	end if
 end if
 rs.close
 conn.close
 set rs=nothing
 set conn=nothing
end if
%>
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="left_usres.asp" -->
<!--#include file="facility.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
        <td background="Images_v1/yx2_2.gif">&nbsp;</td>
        <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
      </tr>
    </table>
        <a href="Product.asp"></a>
        <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td><div align="center"> <img src="Images_v1/lan_gree.gif" width="9" height="19" align="absmiddle"> 会员中心</div></td>
            <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
                <tr>
                  <td></td>
                </tr>
            </table></td>
            <td width="5%"><a href="UserServer.asp"><img src="imgV2/more2.gif" width="37" height="24" border="0"></a></td>
          </tr>
        </table>
      <br>
      <br>
      <%If Session("UserName")="" Then%>
      <table width="100%" border="0" cellpadding="0" cellspacing="3">
        <tr>
          <td colspan="2" align="right"><p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p></td>
        </tr>
        <tr>
          <td colspan="2" align="right"><div align="center">
              <table class="main" cellSpacing="0" cellPadding="2" width="100%" border="0" height="1">
                <tbody>
                  <tr>
                    <td width="100%" height="1"><table width="300" border="0" align="center" cellpadding="0" cellspacing="0" background="Images_v1/bg_4px.gif">
                        <tr>
                          <td colspan="2"><img src="Images_v1/userlogin.gif" width="300" height="11"></td>
                        </tr>
                        <tr>
                          <td width="112"><img src="Images_v1/userlogin8.gif" width="96" height="105" hspace="8"></td>
                          <td width="188"><table border='0' align="center" cellpadding='0' cellspacing='0'>
                              <form action='UserServer.asp?Action=Login' method='post' name='UserLogin' onSubmit='return CheckForm();'>
                                <tr > </tr>
                                <tr>
                                  <td height='25' align='right'>用户名：</td>
                                  <td height='25'><input name='UserName' type='text' id='UserName' size='10' maxlength='20'></td>
                                </tr>
                                <tr>
                                  <td height='25' align='right'>密&nbsp;&nbsp;码：</td>
                                  <td height='25'><input name='Password' type='password' id='Password' size='10' maxlength='20'></td>
                                </tr>
                                <tr align='center'>
                                  <td height='25' colspan='2'><input name='Login' type='submit' id='Login' value=' 登录 '>
                                      <input name='Reset' type='reset' id='Reset' value=' 清除 '></td>
                                </tr>
                                <tr>
                                  <td height='20' colspan='2' align='center'><a href='UserReg.asp' target='_blank'>用户注册</a>&nbsp;&nbsp;<a href='GetPassword.asp' target='_blank'>忘记密码</a></td>
                                </tr>
                              </form>
                          </table></td>
                        </tr>
                        <tr>
                          <td colspan="2"><img src="Images_v1/userlogin2.gif" width="300" height="11"></td>
                        </tr>
                    </table></td>
                  </tr>
                </tbody>
              </table>
          </div></td>
        </tr>
        <tr>
          <td height="19" colspan="2" align="right">&nbsp;</td>
        </tr>
        <tr>
          <td height="19" align="right"><p align="center" style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp;</p></td>
          <td height="19" align="right"><div align="left">请先登录！如果您不是我们的会员，请立即<font color="#006699"><a href="UserReg.asp"><font color="#FF0000">注册</font></a>。</font></div></td>
        </tr>
        <tr>
          <td width="18%" align="right"></td>
          <td width="82%">我们的会员有以下功能：</td>
        </tr>
        <tr>
          <td width="18%" align="right">&nbsp;</td>
          <td width="82%">1、修改您的会员注册资料；</td>
        </tr>
        <tr>
          <td width="18%" align="right">&nbsp;</td>
          <td width="82%">2、专用私人留言簿，您可在此给我们留言和查看我们的回复。</td>
        </tr>
        <tr>
          <td align="right">&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table>
      <%else%>
      <TABLE width="85%" align="center" cellPadding=0 cellSpacing=0>
        <TBODY>
          <TR vAlign=top >
            <TD  width="100%"><p style="margin-bottom: 0"> <br>
              亲爱的<%=Session("UserName")%>：</p>
                <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0">&nbsp; </p>
              <p style="word-spacing: 0; margin-top: 0; margin-bottom: 0"> 　　您现在已经进入会员服务中心，这里只有注册会员才能访问。您可在这里修改您的注册信息、给我们留言、查看我们对您留言的答复。</p></TD>
          </TR>
        </TBODY>
      </TABLE>
      <%end if%></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>
