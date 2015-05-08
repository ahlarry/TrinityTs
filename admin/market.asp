<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="conn.asp"-->
<!--#include file="Admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<%
'管理员删除删除文章-----------
if  Request.QueryString("del")<>"" then
 delsql="delete from market where id="&Request.QueryString("del")
 conn.execute(delsql)
end if
'删除结束---------------

'管理员操作开放/停止标志-----------
if  Request.QueryString("stopflag")<>"" then
    sql="update market set stopflag="&Request.QueryString("stopflag")&" where id="&Request.QueryString("id")
   conn.execute(sql)
end if
'操作开放/停止标志结束---------------

'paixu
if Request("action")="px" then
p="update market set paixu="&Request("paixu")&" where id="&Request("id") 
conn.execute(p)
end if
%>
<script language="JavaScript">
<!--
function cform(){
 if(!confirm("您确认删除吗? 请注意删除后无法恢复!"))
 return false;
}//-->
</script>
<div align="center"> 
  <% If Request.QueryString("province")<>"" Then %>
  <% set rs=Server.Createobject("ADODB.Recordset")
          rs.open "select * from province where id="&Request.QueryString("province"),conn,1,3		 
		  if rs.recordcount=0 then
		    Response.write "没有该省市区"
		  else
		  %>
<table width="96%" border="0" align="center" class="px14">
    <tr> 
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td height="30" align="center" class="back_southidc"><%= rs("province") %> 销 售 网 点 管 理</td>
    </tr>
  </table>
  <% set rs1=Server.Createobject("ADODB.Recordset")
          rs1.open "select * from market where province="&rs("id")&" order by paixu",conn,1,3		 
		  if rs1.recordcount=0 then
		    Response.write "<br><br><br><br><br><center>该地区目前没有销售网点！</center><br><br><br><br>"
		  else
		  %>
  <table width="96%" border="0">
    <tr> 
      <td valign="top">
        <table width="100%" border="0" cellspacing="1">
          <tr bgcolor="#CCCCCC">
            <td width="9%" align="center">排序</td> 
            <td width="47%">网点名称</td>
            <td width="10%" align="center">状态</td>
            <td width="10%" align="center">操作</td>
            <td width="7%" align="center">修改</td>
            <td width="17%" align="center">删除</td>
          </tr>
          <%do while not rs1.eof%>
          <tr bgcolor="#F6F6F6">
		  <form action="" method="post" name="form1">
            <td align="center"><input name="paixu" type="text" class="input" value="<%= rs1("paixu") %>" size="4" maxlength="4">
                <input name="id" type="hidden" id="id" value="<%=rs1("id")%>">
                <input name="action" type="hidden" id="action" value="px">
            </td> 
			</form>
            <td><a href="market_modi.asp?id=<%=rs1("id")%>"><%=rs1("title")%>&nbsp;</a></td>
            <td align="center"> <% If rs1("stopflag")=0 Then %> <font color="#009900">开放</font> <% Else %> <font color="#FF0000">停止</font> <% End If %> </td>
            <td align="center"> <% If rs1("stopflag")=0 Then %> <a href="?province=<%= Request.QueryString("province") %>&stopflag=1&id=<%=rs1("id")%>">停止</a> 
              <% Else %> <a href="?province=<%= Request.QueryString("province") %>&stopflag=0&id=<%=rs1("id")%>">开放</a> 
              <% End If %> </td>
            <td align="center"><a href="market_modi.asp?id=<%=rs1("id")%>">修改</a></td>
            <td align="center"><a href="?province=<%= Request.QueryString("province") %>&del=<%=rs1("id")%>"  onclick="return cform()">删除</a></td>
          </tr>
          <%rs1.movenext
		  loop%>
          <%set rs1=nothing%>
        </table></td>
    </tr>
  </table>
  <% end if  %>
  <% end if  %>
  <% End If %>
<table width="96%"  border="0" align="center" cellpadding="5">
    <tr>
      <td align="center"><img src="../images/chinamap.gif" width="475" height="399" border="0" usemap="#Map4"></td>
    </tr>
  </table>
  <map name="Map4">
    <area shape="poly" coords="273,326,315,310,328,327,307,357,281,351" href="?province=4" alt="广西">
    <area shape="poly" coords="341,317,376,318,382,326,316,366,312,361" href="?province=3" alt="广东">
    <area shape="poly" coords="352,142,370,141,352,162,343,158" href="?province=1" alt="北京">
    <area shape="poly" coords="489,276,487,299" href="#">
    <area shape="poly" coords="389,160,360,157,354,164,363,170,388,170" href="?province=14" alt="天津">
    <area shape="poly" coords="437,238,407,238,406,246,438,248" href="?province=2" alt="上海">
    <area shape="poly" coords="325,374,313,394,302,385,308,376" href="?province=5" alt="海南">
    <area shape="poly" coords="406,285,386,279,368,312,388,321" href="?province=6" alt="福建">
    <area shape="poly" coords="415,255,407,280,389,276,391,253" href="?province=7" alt="浙江">
    <area shape="poly" coords="409,232,382,206,368,215,394,244" href="?province=8" alt="江苏">
    <area shape="poly" coords="383,273,365,311,350,310,350,273,367,266" href="?province=9" alt="江西">
    <area shape="poly" coords="345,276,345,310,306,307,306,274" href="?province=10" alt="湖南">
    <area shape="poly" coords="360,253,356,265,301,272,297,264,305,251,308,232" href="?province=11" alt="湖北">
    <area shape="poly" coords="388,247,384,261,364,258,356,227,367,221" href="?province=12" alt="安徽">
    <area shape="poly" coords="404,177,372,209,352,203,356,182" href="?province=13" alt="山东">
    <area shape="poly" coords="335,199,312,221,318,232,349,244,352,224" href="?province=15" alt="河南">
    <area shape="poly" coords="349,136,329,145,334,196,350,194,354,172,343,158" href="?province=16" alt="河北">
    <area shape="poly" coords="327,158,310,169,310,213,330,200" href="?province=17" alt="山西">
    <area shape="poly" coords="275,248,295,256,297,281,270,280" href="?province=18" alt="重庆">
    <area shape="poly" coords="200,239,270,243,263,284,230,305" href="?province=19" alt="四川">
    <area shape="poly" coords="204,287,200,335,223,360,268,344" href="?province=20" alt="云南">
    <area shape="poly" coords="261,295,294,286,303,309,267,323" href="?province=21" alt="贵州">
    <area shape="poly" coords="406,112,422,131,397,156,372,136" href="?province=22" alt="辽宁">
    <area shape="poly" coords="457,99,429,122,388,87" href="?province=23" alt="吉林">
    <area shape="poly" coords="468,42,452,89,399,77,400,33" href="?province=24" alt="黑龙江">
    <area shape="poly" coords="210,146,281,169,390,110,375,78" href="?province=25" alt="内蒙古">
    <area shape="poly" coords="306,172,304,236,281,241,285,191" href="?province=26" alt="陕西">
    <area shape="poly" coords="189,138,166,162,223,182,237,218,266,235,238,175" href="?province=27" alt="甘肃">
    <area shape="poly" coords="270,175,257,189,270,212,278,200,275,175" href="?province=28" alt="宁夏">
    <area shape="poly" coords="139,176,216,186,228,221,144,235" href="?province=29" alt="青海">
    <area shape="poly" coords="126,63,12,138,32,198,130,190,179,124" href="?province=30" alt="新疆">
    <area shape="poly" coords="40,206,116,204,133,239,188,254,186,275,76,271" href="?province=31" alt="西藏">
    <area shape="poly" coords="361,341,382,346,382,360,359,357,353,347" href="?province=32" alt="香港">
    <area shape="poly" coords="362,369,338,369,342,354,360,360" href="?province=33" alt="澳门">
    <area shape="poly" coords="421,303,426,310,421,338,411,322" href="?province=34" alt="台湾">
  </map>
</div>
<!-- #include file="Inc/Foot.asp" -->