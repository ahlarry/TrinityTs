<%@language=vbscript codepage=936 %>
<%
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
'��Ҫ��ʹ������ֵ�ͼƬ�������
%>
<html>
<head>
<title>����Ա��¼</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/southidc.css">
<script language=javascript>
<!--
function SetFocus()
{
if (document.Login.UserName.value=="")
	document.Login.UserName.focus();
else
	document.Login.UserName.select();
}
function CheckForm()
{
	if(document.Login.UserName.value=="")
	{
		alert("�������û�����");
		document.Login.UserName.focus();
		return false;
	}
	if(document.Login.Password.value == "")
	{
		alert("���������룡");
		document.Login.Password.focus();
		return false;
	}
	if (document.Login.CheckCode.value==""){
       alert ("������������֤�룡");
       document.Login.CheckCode.focus();
       return(false);
    }
}

function CheckBrowser() 
{
  var app=navigator.appName;
  var verStr=navigator.appVersion;
  if (app.indexOf('Netscape') != -1) {
    alert("��������������ʾ��\n    ��ʹ�õ���Netscape����������ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  } 
  else if (app.indexOf('Microsoft') != -1) {
    if (verStr.indexOf("MSIE 3.0")!=-1 || verStr.indexOf("MSIE 4.0") != -1 || verStr.indexOf("MSIE 5.0") != -1 || verStr.indexOf("MSIE 5.1") != -1)
      alert("��������������ʾ��\n    ����������汾̫�ͣ����ܻᵼ���޷�ʹ�ú�̨�Ĳ��ֹ��ܡ�������ʹ�� IE6.0 �����ϰ汾��");
  }
}
//-->
</script>
<style type="text/css">
<!--
.style2 {	COLOR: #003366
}
.style3 {	COLOR: #ffffff
}
.style4 {	FONT-WEIGHT: bold; COLOR: #3d7acd
}
body {
	background-color: #417BC9;
}
.STYLE5 {
	color: #FFFFFF;
	font-weight: bold;
}
.STYLE6 {color: #FFFFFF}
-->
</style>
</head>
<body class="bgcolor">
<p>&nbsp;</p>
<form name="Login" action="Admin_ChkLogin.asp" method="post" target="_parent" onSubmit="return CheckForm();">
    
  <TABLE width=432 height=252 border=0 align="center" cellPadding=0 cellSpacing=0>
    <TBODY>
      <TR>
        <TD align=middle background=image/Login_Top.gif height=26><TABLE cellSpacing=0 cellPadding=0 width=300 border=0>
            <TBODY>
              <TR>
                <TD vAlign=bottom align=middle height=18><span class="STYLE5">��ģ���Բ���վ����ϵͳ</span></TD>
              </TR>
            </TBODY>
        </TABLE></TD>
      </TR>
      <TR>
        <TD align=middle background=image/Login_BG.gif><TABLE height=147 width=400 border=0>
            <TBODY>
              <TR>
                <TD align=right width=147 rowSpan=5><IMG 
                  src="image/Login_TT.jpg" alt="2" width=132 height=130></TD>
                <TD align=right height=30>ϵ��ͳ��</TD>
                <TD width=179 height=30 colspan="2" align=middle><IMG 
                  src="image/Login_BT.gif" alt="1" width=148 height=19> </TD>
              </TR>
              <TR>
                <TD align=right height=30>�û�����</TD>
                <TD height=30 colspan="2" align=middle><input name="UserName"  type="text"  id="UserName4" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></TD>
              </TR>
              <TR>
                <TD align=right height=30>�ܡ��룺</TD>
                <TD height=30 colspan="2" align=middle><input name="Password"  type="password" maxlength="20" style="width:160px;border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); "></TD>
              </TR>
              <TR>
                <TD align=right height=30>��֤�룺</TD>
                <TD height=30 colspan="2" align=middle><input name="CheckCode" size="6" maxlength="4" style="border-style:solid;border-width:1;padding-left:4;padding-right:4;padding-top:1;padding-bottom:1" onMouseOver="this.style.background='#E1F4EE';" onMouseOut="this.style.background='#FFFFFF'" onFocus="this.select(); ">
                <font color="#FFFFFF">�����������<img src="inc/checkcode.asp"></font></TD>
              </TR>
              <TR>
                <TD class=style2 vAlign=center align=right>&nbsp;&nbsp;</TD>
                <TD class=style2 vAlign=center align=right><input   type="submit" name="Submit" value=" ȷ&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'"></TD>
                <TD class=style2 vAlign=center align=right><input name="reset" type="reset"  id="reset" value=" ��&nbsp;�� " style="font-size: 9pt; height: 19; width: 60; color: #000000; background-color: #E1F4EE; border: 1 solid #E1F4EE" onMouseOver ="this.style.backgroundColor='#ffffff'" onMouseOut ="this.style.backgroundColor='#E1F4EE'"></TD>
              </TR>
            </TBODY>
        </TABLE></TD>
      </TR>
      <TR>
        <TD align=middle background=image/Login_Down.gif height=56><TABLE width=400 border=0>
            <TBODY>
              <TR>
                <TD width=160><span class="STYLE6">Ĭ���û�����admin</span></TD>
                <TD width="160"><span class="STYLE6">Ĭ�����룺��������</span></TD>
                <TD><a href="../index.asp"><span class="STYLE6">������ҳ</span></a></TD>
              </TR>
            </TBODY>
        </TABLE></TD>
      </TR>
    </TBODY>
  </TABLE>
  <table width="75%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><div align="center" class="STYLE6"></div></td>
    </tr>
  </table>
  </form>
<script language="JavaScript" type="text/JavaScript">
CheckBrowser();
SetFocus(); 
</script>
</body>
</html>
