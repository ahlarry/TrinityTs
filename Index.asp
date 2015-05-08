<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<%
function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function%>
<style type="text/css">
<!--
.STYLE1 {
	color: #EEEEEE
}
-->
</style>
<!-- #include file="Head.asp" -->
<table width="998" border="0" cellspacing="0" cellpadding="0">
<tr>
<td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
<td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
<td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
  <!--#include file="Left_index.asp" -->
<td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8">
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
    <td background="Images_v1/yx2_2.gif">&nbsp;</td>
    <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="8" cellpadding="2">
  <tr>
    <td valign="top"><table border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="imgV2/news_index1.gif" width="451" height="27"></td>
        </tr>
        <tr>
          <td background="imgV2/news_index2.gif"><table width="100%" border="0" cellspacing="4" cellpadding="0">
              <tr>
                <td width="66%" valign="top"><%
                        set rs_news=server.createobject("adodb.recordset")
                        sqltext4="select top 4 * from news order by id desc"
                        rs_news.open sqltext4,conn,1,1				 	
                        %>
                  <table width="100%" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td width="7%" height="8"></td>
                      <td width="93%"></td>
                    </tr>
                    <%i=0
                          do while not rs_news.eof%>
                    <tr align="center">
                      <td><img src="Img/arrow_2.gif" width="11" height="11"> </td>
                      <td ><div align="left">
                          <p style='line-height:150%'><a href="shownews.asp?id=<%=rs_news("id")%>" target="_blank"> <%=cutstr(rs_news("title"),30)%></a><br>
                        </div></td>
                    </tr>
                    <%rs_news.movenext
						  i=i+1
						  if i=8 then exit do
						  loop
						  rs_news.close %>
                    <tr>
                      <td height="1"></td>
                      <td height="1"></td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td><img src="imgV2/news_index3.gif" width="451" height="17"></td>
        </tr>
      </table></td>
    <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="imgV2/bu_index1.gif" width="246" height="41"></td>
        </tr>
        <tr>
          <td><table width="100%" border="0" cellpadding="0" cellspacing="0" background="imgV2/bu_indexbg.gif">
              <tr>
                <td width="22"><img src="imgV2/bu_index2.gif" width="22" height="135"></td>
                <td valign="top"><table width="93%" border="0" align="center" cellpadding="2" cellspacing="3">
                    <tr>
                      <td><MARQUEE scrollAmount=1 scrollDelay=50 width=200 height="124" direction="up" align="center" onMouseOver="this.stop()" onMouseOut="this.start()">
                        <% call ShowNewBBS() %>
                        </MARQUEE>
                      </td>
                    </tr>
                  </table></td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td><img src="imgV2/bu_index3.gif" width="246" height="52"></td>
        </tr>
      </table></td>
  </tr>
</table>
<table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="16%" colspan="3"><img src="imgV2/index_us.gif" width="114" height="22"></td>
    <td width="79%" valign="middle" colspan="5"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
        <tr>
          <td></td>
        </tr>
      </table></td>
    <td width="5%" colspan="2"><a href="intpld_list.asp"><img src="imgV2/more2.gif" width="37" height="24" border="0"></a></td>
  </tr>
  <tr>
    <td width="1" bgcolor="#4DB5C9">
    <td widtd="3%" class="STYLE4">ID</td>
    <td class="STYLE4">订单号</td>
    <td class="STYLE4">流水号</td>
    <td class="STYLE4">客户名称</td>
    <td class="STYLE4">任务类型</td>
    <td class="STYLE4">调试次数</td>
    <td class="STYLE4">调试结束</td>
    <td class="STYLE4">入库时间</td>
    <td width="1" bgcolor="#4DB5C9">
  </tr>
  <% set re=server.createobject("adodb.recordset")
                set rs = server.CreateObject ("Adodb.recordset") 
sql="select top 5 * from mission where isnull('sjjssj') order by lsh desc" 
rs.open sql,conn,1,1 
			  do while not rs.Eof 
			  %>
  <tr>
    <td width="1" bgcolor="#4DB5C9">
    <td class="STYLE4"><div align="center"><%=i%></div></td>
    <td class="STYLE4"><div align="center"><%=Rs("ddh")%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center"><a href="intpld_dis.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("lsh")%></a></div></td>
    <td class="STYLE4"><div align="center"><%=Rs("khmc")%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center"><%=Rs("rwlx")%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center"><%=Rs("sjtscs")%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("sjjssj"),1)%>&nbsp;</div></td>
    <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("rksj"),1)%>&nbsp;</div></td>
    <td width="1" bgcolor="#4DB5C9">
  </tr>
  <%rs.MoveNext 
loop 
rs.Close 
set rs=nothing 
%>
</table>
<table width="96%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="1" bgcolor="#4DB5C9"></td>
    <td width="724" height="151" valign="top" background="imgV2/index_us2.gif"><table width="98%" border="0" align="center" cellpadding="3" cellspacing="3">
        <tr>
          <td><a href=Aboutus.asp?Title=企业简介 target=_blank>
            <% set rsCompVisualize=server.createobject("adodb.recordset")
			  rsCompVisualize.open "select top 1 * from CompVisualize where ishome=True order by adddate desc",conn,1,1%>
            <% rsCompVisualize.close
				 rsCompVisualize.open "select * from Aboutus where Title='企业简介'",conn,1,1
				 Response.Write(left(rsCompVisualize("content"),360)&"...")%>
            </a>
            <% rsCompVisualize.close
				     set rsCompVisualize=nothing %>
          </td>
        </tr>
      </table></td>
    <td width="1" bgcolor="#4DB5C9"></td>
  </tr>
  <tr>
    <td height="1" colspan="3" bgcolor="#4DB5C9"></td>
  </tr>
  <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
    <br>
    <tr>
      <td width="16%"><img src="imgV2/index_cc.gif" width="114" height="22"></td>
      <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
          <tr>
            <td></td>
          </tr>
        </table></td>
      <td width="5%"><a href="Aboutus.asp"><img src="imgV2/more2.gif" width="37" height="24" border="0"></a></td>
    </tr>
  </table>
  <table width="96%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="1" bgcolor="#4DB5C9"></td>
      <td width="724" height="151" valign="top" background="imgV2/index_us2.gif"><table width="98%" border="0" align="center" cellpadding="3" cellspacing="3">
          <tr>
            <td><a href=Aboutus.asp?Title=企业简介 target=_blank>
              <% set rsCompVisualize=server.createobject("adodb.recordset")
			  rsCompVisualize.open "select top 1 * from CompVisualize where ishome=True order by adddate desc",conn,1,1%>
              <% rsCompVisualize.close
				 rsCompVisualize.open "select * from Aboutus where Title='企业简介'",conn,1,1
				 Response.Write(left(rsCompVisualize("content"),360)&"...")%>
              </a>
              <% rsCompVisualize.close
				     set rsCompVisualize=nothing %>
            </td>
          </tr>
        </table></td>
      <td width="1" bgcolor="#4DB5C9"></td>
    </tr>
    <tr>
      <td height="1" colspan="3" bgcolor="#4DB5C9"></td>
    </tr>
  </table>
  </td>
  
  </tr>
  
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML>