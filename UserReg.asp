<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!-- #include file="Head.asp" -->
<script language=javascript>
function checkreg()
{
  if (document.UserReg.UserName.value=="")
	{
	alert("�������û�����");
	document.UserReg.UserName.focus();
	return false;
	}
  else
    {
	document.reg.username.value=document.UserReg.UserName.value;
	var popupWin = window.open('UserCheckReg.asp', 'CheckReg', 'scrollbars=no,width=340,height=200');
	document.reg.submit();
	}
}
</script>
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
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
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <TR>
          <TD height="1"><FORM name='UserReg' action='UserRegPost.asp' method='post'>
            <br>
            <table width=95% border=0 align="center" cellpadding=5 cellspacing=1 bgcolor="#E4E8EC" style="border-collapse: collapse">
                <TR align=center>
                  <TD height=20 colSpan=2 bgcolor="#F9FBFB"><b>���û�ע��</b></TD>
                </TR>
                <TR>
                  <TD width="32%" bgcolor="#F9FBFB"><b>�û�����</b><BR>
                    ����С��4���ַ���2�����֣�</TD>
                  <TD width="68%" bgcolor="#F9FBFB"><INPUT   maxLength=14 size=30 name=UserName>
                      <font color="#FF0000">*</font>
                      <input name="Check" type="button" id="Check" value="����û���" onClick="checkreg();"></TD>
                </TR>
                <TR>
                  <TD width="32%" bgcolor="#F9FBFB"><B>����(����6λ)��</B><BR>
                    ���ִ�Сд,��Ҫʹ������ '*'��' '�������ַ�</TD>
                  <TD width="68%" bgcolor="#F9FBFB"><INPUT   type=password maxLength=12 size=30 name=Password>
                      <font color="#FF0000">*</font> </TD>
                </TR>
                <TR>
                  <TD width="32%" bgcolor="#F9FBFB"><strong>ȷ������(����6λ)��</strong><BR>
                  </TD>
                  <TD width="68%" bgcolor="#F9FBFB"><INPUT   type=password maxLength=12 size=30 name=PwdConfirm>
                      <font color="#FF0000">*</font> </TD>
                </TR>
                <TR>
                  <TD width="32%" bgcolor="#F9FBFB"><strong>�������⣺</strong><BR>
                    �����������ʾ����</TD>
                  <TD width="68%" bgcolor="#F9FBFB"><INPUT   type=text maxLength=50 size=30 name="Question">
                      <font color="#FF0000">*</font> </TD>
                </TR>
                <TR>
                  <TD width="32%" bgcolor="#F9FBFB"><strong>����𰸣�</strong><BR>
                    �����������ʾ����𰸣�����ȡ������</TD>
                  <TD width="68%" bgcolor="#F9FBFB"><INPUT   type=text maxLength=20 size=30 name="Answer">
                      <font color="#FF0000">*</font> </TD>
                </TR>
              </TABLE>
            <div align="center">
                <INPUT   type=submit value=" ע �� " name=Submit2>
              &nbsp;
              <INPUT name=Reset   type=reset id="Reset" value=" �� �� ">
              </div>
          </form>
              <form name='reg' action='UserCheckreg.asp' method='post' target='CheckReg'>
                <input type='hidden' name='username' value=''>
            </form></TD>
        </TR>
      </table></td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>
