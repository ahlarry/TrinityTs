<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!--#include file="Inc/md5.asp"-->
<!-- #include file="Head.asp" -->
<%
dim UserName,Password,PwdConfirm,Question,Answer,Add
UserName=trim(request("UserName"))
Password=trim(request("Password"))
PwdConfirm=trim(request("PwdConfirm"))
Question=trim(request("Question"))
Answer=trim(request("Answer"))
Add=trim(request("Add"))
if UserName="" or strLength(UserName)>14 or strLength(UserName)<4 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入用户名(不能大于14小于4)</li>"
else
  	if Instr(UserName,"=")>0 or Instr(UserName,"%")>0 or Instr(UserName,chr(32))>0 or Instr(UserName,"?")>0 or Instr(UserName,"&")>0 or Instr(UserName,";")>0 or Instr(UserName,",")>0 or Instr(UserName,"'")>0 or Instr(UserName,",")>0 or Instr(UserName,chr(34))>0 or Instr(UserName,chr(9))>0 or Instr(UserName,"")>0 or Instr(UserName,"$")>0 then
		errmsg=errmsg+"<br><li>用户名中含有非法字符</li>"
		founderr=true
	end if
end if
if Password="" or strLength(Password)>12 or strLength(Password)<6 then
	founderr=true
	errmsg=errmsg & "<br><li>请输入密码(不能大于12小于6)</li>"
else
	if Instr(Password,"=")>0 or Instr(Password,"%")>0 or Instr(Password,chr(32))>0 or Instr(Password,"?")>0 or Instr(Password,"&")>0 or Instr(Password,";")>0 or Instr(Password,",")>0 or Instr(Password,"'")>0 or Instr(Password,",")>0 or Instr(Password,chr(34))>0 or Instr(Password,chr(9))>0 or Instr(Password,"")>0 or Instr(Password,"$")>0 then
		errmsg=errmsg+"<br><li>密码中含有非法字符</li>"
		founderr=true
	end if
end if
if PwdConfirm="" then
	founderr=true
	errmsg=errmsg & "<br><li>请输入确认密码(不能大于12小于6)</li>"
else
	if Password<>PwdConfirm then
		founderr=true
		errmsg=errmsg & "<br><li>密码和确认密码不一致</li>"
	end if
end if
if Question="" then
	founderr=true
	errmsg=errmsg & "<br><li>密码提示问题不能为空</li>"
end if
if Answer="" then
	founderr=true
	errmsg=errmsg & "<br><li>密码答案不能为空</li>"
end if

if founderr=false then
	dim sqlReg,rsReg
	sqlReg="select * from [User] where UserName='" & Username & "'"
	set rsReg=server.createobject("adodb.recordset")
	rsReg.open sqlReg,conn,1,3
	if not(rsReg.bof and rsReg.eof) then
		founderr=true
		errmsg=errmsg & "<br><li>你注册的用户已经存在！请换一个用户名再试试！</li>"
	else
		rsReg.addnew
		rsReg("UserName")=UserName
		rsReg("Password")=md5(Password)
		rsReg("Question")=Question
		rsReg("Answer")=md5(Answer)
		rsReg("Add")=Add
		rsReg("RegDate")=Now()
		rsReg.update
		founderr=false
	end if
	rsReg.close
	set rsReg=nothing
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
            <td><div align="center"> <img src="Images_v1/lan_gree.gif" width="9" height="19" align="absmiddle"> 用户注册</div></td>
            <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
                <tr>
                  <td></td>
                </tr>
            </table></td>
            <td width="5%"><a href="UserServer.asp"><img src="imgV2/more2.gif" width="37" height="24" border="0"></a></td>
          </tr>
        </table>
      <br>
        <table width="96%" border="0" cellspacing="0" cellpadding="0">
          <TR>
            <TD><%
if founderr=false then
	call RegSuccess()
else
	call WriteErrmsg()
end if
%></TD>
          </TR>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>

<%call CloseConn
sub WriteErrMsg()
    response.write "<table align='center' width='60%' border='0' cellpadding='2' cellspacing='5' >"
    response.write "<tr ><td align='center' height='15' bgcolor='#FFCC33'><img src='Images/help.gif' width='24' height='24' align='absmiddle'> 由于以下的原因不能注册用户！</td></tr>"
    response.write "<tr ><td align='left' height='100' bgcolor='#FFFF99'><br>" & errmsg & "<p align='center'>【<a href='javascript:onclick=history.go(-1)'>返 回</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub

sub RegSuccess()
    response.write "<table align='center' width='60%' border='0' cellpadding='2' cellspacing='5' >"
    response.write "<tr ><td align='center' height='15'  bgcolor='#FFCC33'><img src='Images/home.gif' width='20' height='20' align='absmiddle'>成功注册用户！</td></tr>"
    response.write "<tr ><td align='left' height='100'  bgcolor='#FFFF99'><br>你注册的用户名：" & UserName & "<p align='center'>【<a href='javascript:onclick=window.close()'>关 闭</a>】<br></p></td></tr>"
	response.write "</table>" 
end sub
%>
