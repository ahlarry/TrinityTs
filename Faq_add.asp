<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<style type="text/css">
<!--
.STYLE1 {
	color: #FF6666
}
-->
</style>
<!-- #include file="Head.asp" -->
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If
%>
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
      <!--#include file="Left_Faq.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
        </tr>
      </table>
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="imgV2/faq_list.gif" border="0"></td>
          <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
              <tr>
                <td></td>
              </tr>
            </table></td>
          <td width="5%"><img src="imgV2/more2.gif" width="37" height="24" border="0"></td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0" bordercolor="#33FFCC">
        <form id=Faq_add name=Faq_add action=Faq_indb.asp?action=add method=post onSubmit='return checkinf();'>
          <tr>
            <td height="25" colspan="4"></td>
          </tr>        
          <tr>
            <td class=rtd>������:</td>
            <td class=ltd><input type=text name="zrr" size=15>
              <font color="#ff0000">&nbsp;&nbsp;�ɶ��ˣ��ö��ŷָ�</font></td>
          </tr>
          <tr>
            <td class=rtd>�ͻ���:</td>
            <td class=ltd><input type=text name="khm" size=15></td>
          </tr>
          <tr>
            <td class=rtd>��ˮ��:</td>
            <td class=ltd><input type=text name="lsh" size=15></td>
          </tr>
          <tr>
            <td class=rtd>����ʱ��:</td>
            <td class=ltd><script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='sj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script>
            </td>
          </tr>
          <tr>
            <td class=rtd>��Ҫ����:</td>
            <td class=ltd><textarea name="zywt" cols="75" rows="7"></textarea></td>
          </tr>
          <tr> </tr>
          <tr>
            <td class=rtd>ԭ�����:</td>
            <td class=ltd><textarea name="yyfx" cols="75" rows="7"></textarea></td>
          </tr>
          <tr> </tr>
          <tr>
            <td class=rtd>������ʩ:</td>
            <td class=ltd><textarea name="jzcs" cols="75" rows="7"></textarea></td>
          </tr>
          <tr>
            <td height="10" colspan=2>&nbsp;</td>
          </tr>
          <tr>
            <td colspan=2 align="center"><input type=submit value=" �� ȷ �� �� "></td>
          </tr>
        </form>
      </table>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY>
</HTML>
<script language="javascript">
function checkinf(){
	var objdm=document.Faq_add
	if (objdm.zrr.value==""){alert("�����˲���Ϊ��!"); objdm.zrr.focus(); return false;}
	if (objdm.khm.value==""){alert("�ͻ�������Ϊ��!"); objdm.khm.focus(); return false;}
	if (objdm.lsh.value==""){alert("��ˮ�Ų���Ϊ��!"); objdm.lsh.focus(); return false;}
	if (objdm.sj.value==""){alert("ʱ�䲻��Ϊ��!"); objdm.sj.focus(); return false;}
	if (objdm.zywt.value==""){alert("��Ҫ���ⲻ��Ϊ��!");objdm.zywt.focus();return false;}
	if (objdm.yyfx.value==""){alert("ԭ���������Ϊ��!");objdm.yyfx.focus();return false;}
	if (objdm.jzcs.value==""){alert("������ʩ����Ϊ��!");objdm.jzcs.focus();return false;}
	return true;
}
</script>
