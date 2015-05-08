<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If

Dim strFeedBack,Strkhm,strlsh,strzrr,iqiy,iqim,itime,dtstart,dtend
iqiy=Trim(Request("SearchY"))
iqim=Trim(Request("SearchM"))
Strkhm=Trim(Request("khm"))
strzrr=Trim(Request("zrr"))
strlsh=Trim(Request("lsh"))

itime="" : dtstart="" : dtend=""
If iqiy<>"" and iqim<>"" then itime=cdate(iqiy&"年"&iqim&"月1日")
If iqiy<>"" and iqim="" then itime=cdate(iqiy&"年1月1日")
If iqiy="" and iqim<>"" then itime=cdate(year(now)&"年"&iqim&"月1日")
If itime<>"" then
	dtstart=itime
	dtend=dateadd("m",1,dtstart)
	dtend=dateadd("d",-1,dtend)
End If

strFeedBack="&lsh="&strlsh&"&khm="&strkhm&"&zrr="&strzrr&"&SearchY="&iqiy&"&SearchM="&iqim
%>
<style type="text/css">
<!--
.STYLE2 {
	font-size: x-large;
	font-weight: bold;
	color: #000066;
}
.STYLE4 {
	color: #000066;
}
.btn_2k3 {
	BORDER-RIGHT: #002D96 1px solid;
	PADDING-RIGHT: 2px;
	BORDER-TOP: #002D96 1px solid;
	PADDING-LEFT: 2px;
	FONT-SIZE: 12px;
FILTER: progid:DXImageTransform.Microsoft.Gradient(GradientType=0, StartColorStr=#FFFFFF, EndColorStr=#9DBCEA);
	BORDER-LEFT: #002D96 1px solid;
	CURSOR: hand;
	COLOR: black;
	PADDING-TOP: 2px;
	BORDER-BOTTOM: #002D96 1px solid
}
-->
</style>
<!-- #include file="Head.asp" -->
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
      <%Call SearchInfo()%>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" colspan="2" bgcolor="#00ffff"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </table>
      <%     
strSql=""
If strlsh <> "" Then
	If strsql<>"" Then 
		strSql = strSql&"and  lsh like '%"&strlsh&"%' "
	else
		strSql="lsh like '%"&strlsh&"%' "
	End If
End If	
If strkhm <> "" Then
	If strsql<>"" Then 
		strSql = strSql&"and  Khmc like '%"&strkhm&"%' "
	else
		strSql="Khmc like '%"&strkhm&"%' "
	End If
End If	

If dtstart<>"" then
	If strSql="" then
		strSql="datediff('d',sj,'"&dtstart&"')<=0 and datediff('d',sj,'"&dtend&"')>=0"
	else
		strSql=strSql&"and datediff('d',sj,'"&dtstart&"')<=0 and datediff('d',sj,'"&dtend&"')>=0"
	End If
End If
		
set rs=server.createobject("adodb.recordset")
If strSql="" then
	sqltex="select * from Faq Order BY sj desc"
else
	sqltex="select * from Faq where "&strSql&" Order BY sj desc"
End If
rs.open sqltex,conn,1,1
dim PerPage
PerPage=30
'假如没有数据时
If rs.eof and rs.bof then
response.write ("<p align='center'><font color='#ff0000'>当前条件下没有问题！</font></p>")
else
'取得页数,并判断用户输入的是否数字类型的数据，如不是将以第一页显示
text="0123456789"
Rs.PageSize=PerPage
for i=1 to len(request("page"))
checkpage=instr(1,text,mid(request("page"),i,1))
if checkpage=0 then
exit for
end if
next

If checkpage<>0 then
  If NOT IsEmpty(request("page")) Then
   CurrentPage=Cint(request("page"))
   If CurrentPage < 1 Then CurrentPage = 1 
   If CurrentPage > Rs.PageCount Then CurrentPage = Rs.PageCount 
  Else
   CurrentPage= 1
  End If
   If not Rs.eof Then Rs.AbsolutePage = CurrentPage end if
Else
 CurrentPage=1
End if
  %>
      <table width="98%" border="1" cellspacing="0"  bordercolor="#33FFCC">
        <caption>
        <span class="STYLE2">挤出模厂内调试问题分析列表</span> <br />
        </caption>
        <tr>
          <th width="3%" class="STYLE4">ID</th>
          <th class="STYLE4">流水号</th>
          <th class="STYLE4">客户名称</th>
          <th class="STYLE4">责任人</th>
          <th class="STYLE4">主要问题</th>
          <th class="STYLE4">解决措施</th>
          <th class="STYLE4">时间</th>
        </tr>
        <%if not rs.eof then
i=1
do while not rs.eof
%>
        <tr>
          <td class="STYLE4"><div align="center"><%=i%></div></td>
          <td class="STYLE4"><div align="center"><a href="Faq_dis.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("lsh")%></a></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("khmc")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=Rs("zrr")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=Rs("WenTi")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=Rs("CuoS")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("sj"),1)%>&nbsp;</div></td>
        </tr>
        <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
      </table>
      <%
      Call showpage("Faq.asp?"&strFeedBack,Rs.RecordCount,PerPage,True,True,"条")
      End If%>
    </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML><%
Function SearchInfo()
'Response.Write("iqiy的值为"&iqiy&",数据类型为"&vartype(iqiy)&"<br>")
'Response.Write("iqim的值为"&iqim&",数据类型为"&vartype(iqim)&"<br>")
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
          <tr>
            <td>&nbsp;&nbsp;筛选条件:</td>
            <td>时间:<select size="1" id="SearchY" name="SearchY">
            	<option value="" selected="selected"></option>
                  <%for i=year(now)-2 to year(now)+2%>
                  <option value=<%=i%> <%if iqiy = i Then%>selected<%End If%>><%=i%></option>
                  <%next%>
            </select>年
            <select size="1" id="SearchM" name="SearchM">
            	<option value="" selected="selected"></option>
                <%for i=1 to 12%>
          		      <option value="<%=i%>" <%if iqim=i Then%>selected<%End If%>><%=i%></option>
                <%next%>      
            </select>月
              </td>      
            <td>责任人:
              <input name="zrr" size="10" value="<%=strzrr%>"></td>
            <td>客户名称:
              <input name="khm" size="15" value="<%=strkhm%>">
            </td>
            <td>流水号:
              <input name="lsh" size="10" value="<%=strlsh%>"></td>
            <td align="left"><input class=btn_2k3 type="submit" value="检 索" /></td>
          </tr>
  </form>
</table>
<%
End Function
%>
