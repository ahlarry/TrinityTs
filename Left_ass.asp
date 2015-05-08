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
          <TD width="100%"><a href='assess_add.asp'>添加考评</a></TD>
        </tr>
        <tr>
          <TD >&nbsp;</TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%"><a href='ass_list.asp'>考评列表 </a></TD>
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
          <TD width="100%">流水号：
            <input name="czlsh" type="text" id="czlsh" size="15" />
            </a></TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">订单号：
            <input name="czddh" type="text" id="czddh" size="15" />
            </a></TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">客户名：
            <input name="czkhm" type="text" id="czkhm" size="15" />
            </a></TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">生产线：
            <input name="czscx" type="text" id="czscx" size="15" />
          </TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">调试员/操作工：
            <input name="czczg" type="text" id="czczg" size="8" />
          </TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">从
            <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='czkssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script>
            </a> </TD>
        </tr>
        <tr>
          <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
          <TD width="100%">到
            <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='czjssj';  //注意这里设置输入框的name，同一页中日期输入框，不能出现重复的name。
  		myDate.display();
		</script>
            </a> </TD>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <TD width=34 height=24 align=center>&nbsp;</TD>
          <TD width="100%"><input type="submit" name="Submit"  value="搜  索">
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
