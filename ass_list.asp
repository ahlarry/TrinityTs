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

Dim strFeedBack,Strkhm,strlsh,strzrr,iqiy,iqim,itime,dtstart,dtend
iqiy=Trim(Request("SearchY"))
iqim=Trim(Request("SearchM"))
strzrr=Trim(Request("zrr"))
strkpnr=Trim(Request("kpnr"))

itime="" : dtstart="" : dtend=""
If iqiy<>"" and iqim<>"" then itime=cdate(iqiy&"年"&iqim&"月1日")
If iqiy<>"" and iqim="" then itime=cdate(iqiy&"年1月1日")
If iqiy="" and iqim<>"" then itime=cdate(year(now)&"年"&iqim&"月1日")
If itime<>"" then
	dtstart=itime
	dtend=dateadd("m",1,dtstart)
	dtend=dateadd("d",-1,dtend)
End If

strFeedBack="&kpnr="&strkpnr&"&zrr="&strzrr&"&SearchY="&iqiy&"&SearchM="&iqim
%>
<table width="998" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
    <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510"></td>
    <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
      <!--#include file="Left_ass.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
        </tr>
      </table>
      <table width="96%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="imgV2/ass_list.gif" border="0"></td>
          <td width="79%" valign="middle"><table width="100%" height="1" border="0" cellpadding="0" cellspacing="0" bgcolor="#00AD97">
              <tr>
                <td></td>
              </tr>
            </table></td>
          <td width="5%"><img src="imgV2/more2.gif" width="37" height="24" border="0"></td>
        </tr>
      </table>

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
            <td>考评内容:
              <input name="kpnr" size="15" value="<%=strkpnr%>"></td>
            <td align="left"><input class=btn_2k3 type="submit" value="检 索" /></td>
          </tr>
  </form>
        <tr>
          <td height="1" colspan="6">&nbsp;</td>
        </tr>
        <tr>
          <td height="1" colspan="6" bgcolor="#00ffff"></td>
        </tr>
        <tr>
          <td height="1" colspan="6">&nbsp;</td>
        </tr>
      </table>
   
      <%     
strSql=""
If strkpnr <> "" Then
	If strsql<>"" Then 
		strSql = strSql&"and  kpnr like '%"&strkpnr&"%' "
	else
		strSql="kpnr like '%"&strkpnr&"%' "
	End If
End If	
If strzrr <> "" Then
	If strsql<>"" Then 
		strSql = strSql&"and  zrr like '%"&strzrr&"%' "
	else
		strSql="zrr like '%"&strzrr&"%' "
	End If
End If	

If dtstart<>"" then
	If strSql="" then
		strSql="datediff('d',kpsj,'"&dtstart&"')<=0 and datediff('d',kpsj,'"&dtend&"')>=0"
	else
		strSql=strSql&"and datediff('d',kpsj,'"&dtstart&"')<=0 and datediff('d',kpsj,'"&dtend&"')>=0"
	End If
End If
		
set rs=server.createobject("adodb.recordset")
If strSql="" then
	sqltex="select * from kplb Order BY kpsj desc"
else
	sqltex="select * from kplb where "&strSql&" Order BY kpsj desc"
End If
rs.open sqltex,conn,1,1
dim PerPage
PerPage=30
'假如没有数据时
If rs.eof and rs.bof then
response.write "<p align='center'><font color='#ff0000'>还没有任务！</font></p>"
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
      <table width="98%" border="2" bordercolor="#33FFCC">
        <tr>
          <th width="3%" class="STYLE4">ID</th>
          <th width="10%" class="STYLE4">责任人</th>
          <th class="STYLE4">考评内容</th>
          <th width="10%" class="STYLE4">考评分</th>
          <th width="13%" class="STYLE4">考评时间</th>
          <th width="10%" class="STYLE4">考评人</th>
          <th width="7%" class="STYLE4">操作</th>
        </tr>
        <%if not rs.eof then
i=1
do while not rs.eof
%>
        <tr>
          <td class="STYLE4"><div align="center"><%=i%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("zrr")%></div></td>
          <td class="STYLE4"><div align="center" Title="备注:<%=Rs("bz")%>"><%=Rs("kpnr")%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("fz")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=Rs("kpsj")%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("kpr")%></div></td>
          <td class="STYLE4"><div align="center">
            <%If session("GZZ")<=0 Then%>
            <a href="ass_indb.asp?action=del&id=<%=Rs("ID")%>" onClick="return confirm('删除后将不能恢复!\n\n确认删除?');">删除</a>
            <%else%>
            &nbsp;
            <%End If%></td>
        </tr>
        <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
End If
%>
      </table>
      <%
      Call showpage("ass_list.asp?",Rs.RecordCount,PerPage,True,True,"条")
      %>
</table>
<!-- #include file="Inc/Foot.asp" -->
</BODY>
</HTML>