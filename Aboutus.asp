<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_css.asp"-->
<!-- #include file="Head.asp" -->
<%Title=Trim(request("Title"))
Set rs = Server.CreateObject("ADODB.Recordset")
sql="select Content from Aboutus where Title='"&Title&"'"
rs.open sql,conn,1,3
%>
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><img src="imgV2/left_us.gif" width="211" height="30"></td>
          </tr>
          <tr>
            <td><table width="80%" border="0" align="center" cellpadding="0" cellspacing="0">
              <%
Set rslist = Server.CreateObject("ADODB.Recordset")
sql="select Title from Aboutus order by Aboutusorder"
rslist.open sql,conn,1,3
do while not rslist.eof
%>
              <tr>
                <td height="25"><div align="center"> <img src="Img/arrow_4.gif" width="12" height="11" align="absmiddle"> <a href="Aboutus.asp?Title=<%=rslist("Title")%>"><%=rslist("Title")%></a></div></td>
              </tr>
              <TR>
                <TD 
                            height=1 colspan="2" 
                            background=img/naSzarym.gif><IMG height=1 src="img/1x1_pix.gif" 
                              width=10></TD>
              </TR>
              <%rslist.movenext 
loop%>
              <tr>
                <td height="25"><div align="center"><img src="Img/arrow_4.gif" width="12" height="11" align="absmiddle"> <a href="Culture.asp">企业文化</a></div></td>
              </tr>
              <TR>
                <TD 
                            height=1 colspan="2" 
                            background=img/naSzarym.gif><IMG height=1 src="img/1x1_pix.gif" 
                              width=10></TD>
              </TR>
            </table></td>
          </tr>
        </table>
      <!--#include file="left_product.asp" -->
      <table width="100%" height="35" border="0" cellpadding="0" cellspacing="0" background="imgV2/left_frisbg.gif">
          <tr>
            <td align="center" valign="middle"><% call ShowFriendLinks(2,6,6,3) %></td>
          </tr>
      </table></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
        <td background="Images_v1/yx2_2.gif">&nbsp;</td>
        <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
      </tr>
    </table>
        <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="16%"><a href="Product.asp"><img src="imgV2/abouts.gif" width="115" height="23" border="0"></a></td>
            <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
                <tr>
                  <td></td>
                </tr>
            </table></td>
            <td width="5%"><a href="Aboutus.asp?Title=企业简介"><img src="imgV2/more2.gif" width="37" height="24" border="0"></a></td>
          </tr>
        </table>
      <a href="Product.asp"></a>
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FAFBFB">
        <tr>
          <td width="14" background="Images_v1/yx2_4.gif">&nbsp;</td>
          <td bgcolor="#F8FFE8"><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><table width="98%" border="0" align="right" cellpadding="0" cellspacing="0" background="Images_v1/Prou_bg.gif">
                    <tr>
                      <td width="30%"><img src="Images_v1/i6.gif" width="13" height="13" hspace="5" align="absmiddle"><strong> <%=Title%></strong></td>
                      <td width="63%">&nbsp;</td>
                      <td width="7%" align="right"><img src="Images_v1/Prou2.gif" width="11" height="30"></td>
                    </tr>
                </table></td>
              </tr>
              <tr>
                <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="3%" height="34" rowspan="3">&nbsp;</td>
                      <td width="94%">&nbsp;</td>
                      <td width="3%" rowspan="3">&nbsp;</td>
                    </tr>
                    <tr>
                      <td><%=rs("Content")%></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                </table></td>
              </tr>
          </table></td>
          <td width="21" align="right" background="Images_v1/yx2_5.gif">&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
<%rs.close
set rs=nothing
rslist.close
set rslist=nothing
%>
<!-- #include file="Inc/Foot.asp" -->
</BODY></HTML>
