<!--#include file="inc/calendar.asp"-->

<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="imgV2/left_ass.gif" width="211" height="30" /></td>
  </tr>
  <tr>
    <td align="center"><table width="86%" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <TD >&nbsp;</TD>
        </tr>    
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%"><a href='assess_add.asp'>��ӿ���</a></TD>
        </tr>
        <tr>
          <TD >&nbsp;</TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%"><a href='ass_list.asp'>�����б� </a></TD>
        </tr>        
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="imgV2/left_so2.gif" width="211" height="30" /></td>
  </tr>
  <tr>
    <td align="center"><table width="86%" border="0" align="center" cellpadding="0" cellspacing="2">
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">��ˮ�ţ�
            <input name="czlsh" type="text" id="czlsh" size="15" />
            </a></TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">�����ţ�
            <input name="czddh" type="text" id="czddh" size="15" />
            </a></TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">�ͻ�����
            <input name="czkhm" type="text" id="czkhm" size="15" />
            </a></TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">�����ߣ�
            <input name="czscx" type="text" id="czscx" size="15" />
          </TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">����Ա/��������
            <input name="czczg" type="text" id="czczg" size="8" />
          </TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">��
            <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='czkssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script>
            </a> </TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">��
            <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='czjssj';  //ע����������������name��ͬһҳ����������򣬲��ܳ����ظ���name��
  		myDate.display();
		</script>
            </a> </TD>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <TD width=34 height=24 align=center>&nbsp;</TD>
          <TD width="100%"><input type="submit" name="Submit"  value="��  ��">
            </a></TD>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="100%" height="35" border="0" cellpadding="0" cellspacing="0" background="imgV2/left_frisbg.gif">
  <tr>
    <td align="center" valign="middle"><% call ShowFriendLinks(2,6,6,3) %></td>
  </tr>
</table>
