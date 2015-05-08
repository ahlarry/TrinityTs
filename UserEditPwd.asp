<!--#include file="inc/conn.asp"-->
<!--#include file="inc/Skin_css.asp"-->
<!--#include file="inc/config.asp"-->
<!--#include file="inc/md5.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim Action,UserName,FoundErr,ErrMsg
dim rsUser,sqlUser
Action=trim(request("Action"))
UserName=trim(request("UserName"))
if Action="" and session("UserName")="" then
	response.redirect "UserServer.asp"
end if
if Action="Modify" and UserName<>"" then
	Set rsUser=Server.CreateObject("Adodb.RecordSet")
	sqlUser="select * from [User] where UserName='" & UserName & "'"
	rsUser.Open sqlUser,conn,1,3
	if rsUser.bof and rsUser.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的用户！</li>"
	else
		dim OldPassword,Password,PwdConfirm
		OldPassword=trim(request("OldPassword"))
		Password=trim(request("Password"))
		PwdConfirm=trim(request("PwdConfirm"))
		if OldPassword="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入旧密码！</li>"
		else
			if Instr(OldPassword,"=")>0 or Instr(OldPassword,"%")>0 or Instr(OldPassword,chr(32))>0 or Instr(OldPassword,"?")>0 or Instr(OldPassword,"&")>0 or Instr(OldPassword,";")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,"'")>0 or Instr(OldPassword,",")>0 or Instr(OldPassword,chr(34))>0 or Instr(OldPassword,chr(9))>0 or Instr(OldPassword,"")>0 or Instr(OldPassword,"$")>0 then
				errmsg=errmsg+"<br><li>旧密码中含有非法字符</li>"
				founderr=true
			else
				if md5(OldPassword)<>rsUser("Password") then
					FoundErr=True
					ErrMsg=ErrMsg & "<br><li>你输入的旧密码不对，没有权限修改！</li>"
				end if
			end if
		end if
		if strLength(Password)>12 or strLength(Password)<6 then
			founderr=true
			errmsg=errmsg & "<br><li>请输入新密码(不能大于12小于6)。</li>"
		else
			if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
				errmsg=errmsg+"<br><li>新密码中含有非法字符</li>"
				founderr=true
			end if
		end if
		if PwdConfirm="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请输入确认密码！</li>"
		else
			if PwdConfirm<>Password then
				FoundErr=True
				ErrMsg=ErrMsg & "<br><li>确认密码与新密码不一致！</li>"
			end if
		end if
		if FoundErr<>true then
			rsUser("Password")=md5(Password)
			rsUser.update
		end if
	end if
	rsUser.close
	set rsUser=nothing
	if FoundErr=True then
		call WriteErrMsg()
	else
	    response.write"<SCRIPT language=JavaScript>alert('成功修改密码！');"
        response.write"javascript:history.go(-1)</SCRIPT>"
	end if
else
%>
<!-- #include file="Head.asp" -->
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
        <p>&nbsp;</p>
        <table width="96%" border="0" cellspacing="0" cellpadding="0">
          <TR>
            <TD><FORM name="Form1" action="UserEditPwd.asp" method="post">
                  <table width=90% border=0 align="center" cellpadding=2 cellspacing=0 class='border'>
                    <TR align=center >
                      <TD width="120" height=7 bgcolor="#FB9D24"><img src="Images_v1/UserEdit.gif" width="228" height="25" vspace="2"></TD>
                    </TR>
                    <TR align=center >
                      <TD height=1 bgcolor="#CC3300"></TD>
                    </TR>
                    <TR align=center >
                      <TD height=1><table width=400 border=0 align="center" cellpadding=2 cellspacing=2>
                          <TR align=center > </TR>
                          <TR>
                            <TD width="120" align="right">用 户 名：</TD>
                            <TD><%=session("UserName")%>
                                <input name="UserName" type="hidden" value="<%=session("UserName")%>"></TD>
                          </TR>
                          <TR>
                            <TD width="120" align="right">旧 密 码：</TD>
                            <TD><INPUT   type=password maxLength=16 size=30 name=OldPassword>
                            </TD>
                          </TR>
                          <TR>
                            <TD width="120" align="right">新 密 码：</TD>
                            <TD><INPUT   type=password maxLength=16 size=30 name=Password>
                            </TD>
                          </TR>
                          <TR>
                            <TD width="120" align="right">确认密码：</TD>
                            <TD><INPUT name=PwdConfirm   type=password id="PwdConfirm" size=30 maxLength=16>
                            </TD>
                          </TR>
                          <TR align="center">
                            <TD height="40" colspan="2"><input name="Action" type="hidden" id="Action" value="Modify">
                                <input name=Submit   type=submit id="Submit" value=" 保 存 ">
                            </TD>
                          </TR>
                      </TABLE></TD>
                    </TR>
                  </TABLE>
                </form></TD>
          </TR>
        </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>
<%
end if
call CloseConn()
%>