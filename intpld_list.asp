<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If

function cutstr(tempstr,tempwid)
if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&"..."
else
cutstr=tempstr
end if
end function

Dim strFeedBack,StrRwlx,Strkhm,strddh,strlsh,strrwzt,strrwlr,iqiy,iqim,izhiy,izhim,iqitime,izhitime,dtstart,dtend
StrRwlx=Request("rwlx")
Strrwlr=Request("rwlr")
iqiy=Trim(Request("qiy"))
iqim=Trim(Request("qim"))
izhiy=Trim(Request("zhiy"))
izhim=Trim(Request("zhim"))
Strkhm=Trim(Request("khm"))
strddh=Trim(Request("ddh"))
strlsh=Trim(Request("lsh"))
strrwzt=Request("rwzt")
iqitime="" : izhitime="" : dtstart="" : dtend=""
If iqiy<>"" and iqim<>"" then iqitime=cdate(iqiy&"年"&iqim&"月1日")
If izhiy<>"" and izhim<>"" then izhitime=cdate(izhiy&"年"&izhim&"月1日")
If iqitime<>"" and izhitime="" then
	dtstart=iqitime
	dtend=dateadd("m",1,dtstart)
	dtend=dateadd("d",-1,dtend)
End If
If iqitime="" and izhitime<>"" then
	dtstart=izhitime
	dtend=dateadd("m",1,dtstart)
	dtend=dateadd("d",-1,dtend)
End If
If iqitime<>"" and izhitime<>"" then
	If datediff("d",iqitime,izhitime)>=0 then
		dtstart=iqitime
		dtend=dateadd("m",1,izhitime)
		dtend=dateadd("d",-1,dtend)
	else
		dtstart=izhitime
		dtend=dateadd("m",1,iqitime)
		dtend=dateadd("d",-1,dtend)
	End If
End If

If StrRwlx="" Then StrRwlx="1" End If
strFeedBack="rwlx="&StrRwlx&"&ddh="&strddh&"&khm="&strkhm&"&rwzt="&strrwzt&"&qiy="&iqiy&"&qim="&iqim&"&zhiy="&izhiy&"&zhim="&izhim&"&rwlr="&strrwlr
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
      <!--#include file="Left_pld.asp" --></td>
    <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21"></td>
          <td background="Images_v1/yx2_2.gif">&nbsp;</td>
          <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21"></td>
        </tr>
      </table>
      <table width="98%" border="0" cellspacing="0" cellpadding="0">
        <tbody>
          <tr>
            <td width="40%" height="18">　计划任务&gt;&gt; 调试任务列表 </td>
            <td width="71%"><div align="right">
                <label>
                <input name="rwlx" type="radio" value="1" onClick="location.href='intpld_list.asp?rwlx='+ this.value" <%If strrwlx="1" Then%>="" checked="checked" <%End If%>="" />
                厂内调试
                <input name="rwlx" type="radio" value="0" onClick="location.href='intpld_list.asp?rwlx='+ this.value" <%If strrwlx="0" Then%>="" checked="checked" <%End If%>="" />
                厂外调试</label>
              </div></td>
          </tr>
          <tr>
            <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
          </tr>
          <tr>
            <td height="1" colspan="2">&nbsp;</td>
          </tr>
        </tbody>
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
If strddh <> "" Then
	strSql = "ddh like '%"&strddh&"%' "
End If
If strlsh <> "" Then
	If strsql<>"" Then
		strSql = strSql&"and  lsh like '%"&strlsh&"%' "
	else
		strSql="lsh like '%"&strlsh&"%' "
	End If
End If
If strkhm <> "" Then
	If strsql<>"" Then
		strSql = strSql&"and  khmc like '%"&strkhm&"%' "
	else
		strSql="khmc like '%"&strkhm&"%' "
	End If
End If
Select case strrwlr
	case "1"
		If strSql="" then
			strSql = "rwlx='全套' "
		else
			strSql=strSql&"and rwlx='全套' "
		End If
	case "2"
		If strSql="" then
			strSql = "rwlx='机头' "
		else
			strSql=strSql&"and rwlx='机头' "
		End If
	case "3"
		If strSql="" then
			strSql = "rwlx='定型' "
		else
			strSql=strSql&"and rwlx='定型' "
		End If
	case "4"
		If strSql="" then
			strSql = "rwlx='新品' "
		else
			strSql=strSql&"and rwlx='新品' "
		End If
	case "5"
		If strSql="" then
			strSql = "rwlx='非调' "
		else
			strSql=strSql&"and rwlx='非调' "
		End If
	case "6"
		If strSql="" then
			strSql = "rwlx='TB' "
		else
			strSql=strSql&"and rwlx='TB' "
		End If
End Select
Select case strrwzt
	case "0"
		If StrRwlx=1 Then
			iqitime=""
		else
			iqitime="cksj"
		End If
	case "1"
		If StrRwlx=1 Then
			iqitime=""
			If strSql="" then
				strSql = "isNull(sjjssj) and InStr('非调TB',rwlx)=0"
			else
				strSql=strSql&"and isNull(sjjssj) and InStr('非调TB',rwlx)=0"
			End If
		else
			iqitime="cksj"
			If strSql="" then
				strSql = "not(isNull(cksj)) and isNull(cwjssj) and isNull(btjsj)"
			else
				strSql=strSql&"and not(isNull(cksj)) and isNull(cwjssj) and isNull(btjsj)"
			End If
		End If
	case "2"
		If StrRwlx=1 Then
			iqitime="sjjssj"
			If strSql="" then
				strSql = "(not(isNull(sjjssj)) and isNull(rksj)) or (InStr('非调TB',rwlx)>0 and isNull(rksj))"
			else
				strSql=strSql&"and ((not(isNull(sjjssj)) and isNull(rksj)) or (InStr('非调TB',rwlx)>0 and isNull(rksj))) "
			End If
		else
			iqitime="cwjssj"
			If strSql="" then
				strSql = "(not(isNull(cwjssj))or not(isNull(btjsj)))"
			else
				strSql=strSql&"and (not(isNull(cwjssj))or not(isNull(btjsj)))"
			End If
		End If
	case "3"
		iqitime="rksj"
		If strSql="" then
			strSql = "not(isNull(rksj)) and isNull(cksj)"
		else
			strSql=strSql&"and not(isNull(rksj)) and isNull(cksj)"
		End If
	case "4"
		iqitime="cksj"
		If strSql="" then
			strSql = "not(isNull(cksj)) "
		else
			strSql=strSql&"and not(isNull(cksj)) "
		End If

End Select

'If StrRwlx="0" Then iqitime="cwjssj"
If dtstart<>"" then
	If iqitime<>"" then
		If strSql="" then
			strSql="datediff('d',"&iqitime&",'"&dtstart&"')<=0 and datediff('d',"&iqitime&",'"&dtend&"')>=0"
		else
			strSql=strSql&"and datediff('d',"&iqitime&",'"&dtstart&"')<=0 and datediff('d',"&iqitime&",'"&dtend&"')>=0"
		End If
	End If
End If

set rs=server.createobject("adodb.recordset")
If StrRwlx="1" Then
	If strSql="" then
		sqltex="select * from mission Order BY ddh,lsh desc"
	else
		sqltex="select * from mission where "&strSql&" Order BY ddh,lsh desc"
	End If
else
	If strSql="" then
		sqltex="select * from mission where not(isNull(cksj)) Order BY cwkssj desc,lsh desc"
	else
		sqltex="select * from mission where "&strSql&" and not(isNull(cksj)) Order BY cwkssj desc,lsh desc"
	End If
End If
rs.open sqltex,conn,1,1
dim PerPage
PerPage=20
'假如没有数据时
If rs.eof and rs.bof then
response.write ("<p align='center'><font color='#ff0000'>还没有任务！"&dtstart&"-"&dtend&"</font></p>")
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
If StrRwlx="1" Then
  %>
      <table width="98%" border="1" cellspacing="0"  bordercolor="#33FFCC">
        <caption>
        <span class="STYLE2">挤出模厂内调试任务列表</span> <br />
        </caption>
        <tr>
          <th width="3%" class="STYLE4">ID</th>
          <th class="STYLE4">订单号</th>
          <th class="STYLE4">流水号</th>
          <th class="STYLE4">客户名</th>
          <th class="STYLE4">任务</th>
          <th class="STYLE4">调试次数</th>
          <th class="STYLE4">调试结束</th>
          <th class="STYLE4">入库时间</th>
          <th class="STYLE4">出库时间</th>
          <th class="STYLE4">操作</th>
        </tr>
        <%if not rs.eof then
i=1
do while not rs.eof
%>
        <tr>
          <td class="STYLE4"><div align="center"><%=i%></div></td>
          <td class="STYLE4"><div align="center"><%=Rs("ddh")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><a href="intpld_dis.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("lsh")%></a></div></td>
          <td class="STYLE4"><div align="center"><%=cutstr(Rs("khmc"),10)%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=Rs("rwlx")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=Rs("sjtscs")%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("sjjssj"),2)%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("rksj"),2)%>&nbsp;</div></td>
          <td class="STYLE4"><div align="center"><%=FormatDateTime(Rs("cksj"),2)%>&nbsp;</div></td>
          <td class="STYLE4"><label>
            <div align="center"> <a href="pld_change.asp?s_lsh=<%=Rs("lsh")%>">更改</a>|<a href="intpld_del.asp?id=<%=Rs("ID")%>&lsh=<%=Rs("lsh")%>" onClick="return confirm('删除后将不能恢复!\n\n确认删除  <%=Rs("lsh")%> 任务书?');">删除</a>
              </label>
            </div></td>
        </tr>
        <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
      </table>
      <%else%>
      <table width="98%" border="1" cellspacing="0" bordercolor="#33FFCC">
        <caption>
        <span class="STYLE2">挤出模厂外调试任务列表</span> <br />
        </caption>
        <tr>
          <th scope="col">ID</th>
          <th scope="col">流水号</th>
          <th scope="col">客户名</th>
          <th scope="col">出库时间</th>
          <th scope="col">调试开始</th>
          <th scope="col">调试结束</th>
          <th scope="col">验收结果</th>
          <th scope="col">操作</th>
        </tr>
        <%if not rs.eof then
i=1
do while not rs.eof
%>
        <tr>
          <td><div align="center"><%=i%></div></td>
          <td><div align="center"><a href="intpld_dis.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("lsh")%>&nbsp;</a></div></td>
          <td><div align="center"><%=Rs("khmc")%>&nbsp;</div></td>
          <td><div align="center"><%=Rs("cksj")%>&nbsp;</div></td>
          <td><div align="center"><%=Rs("cwkssj")%>&nbsp;</div></td>
          <td><div align="center"><%If isNull(Rs("btjsj")) Then
          								Response.Write(Rs("cwjssj"))
          							else
          								Response.Write(Rs("btjsj")&"�")
          							End If%>&nbsp;</div></td>
          <td><div align="center" alt="<%=Rs("bz")%>"><%=Rs("ysjg")%>&nbsp;</div></td>
          <td><label>
            <div align="center"> <a href="extpld_change.asp?id=<%=Rs("ID")%>">更改</a> | <a href="extpld_del.asp?id=<%=Rs("ID")%>" onClick="return confirm('删除后将不能恢复!\n\n确认删除  <%=Rs("lsh")%> 任务书?');">删除</a>
              </label>
            </div></td>
        </tr>
        <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
      </table>
      <%End If
      Call showpage("intpld_list.asp?"&strFeedBack,Rs.RecordCount,PerPage,True,True,"条")
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
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td rowspan="2">&nbsp;&nbsp;筛选条件:</td>
      <td><select name="qiy">
          <option value=""></option>
          <%for i=year(now)-2 to year(now)+2%>
          <option value="<%=i%>" <%If iqiy=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="qim">
          <option value=""></option>
          <%for i=1 to 12%>
          <option value="<%=i%>" <%If iqim=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        月&nbsp;-&nbsp;
        <select name="zhiy">
          <option value=""></option>
          <%for i=year(now)-2 to year(now)+2%>
          <option value="<%=i%>" <%If izhiy=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        年
        <select name="zhim">
          <option value=""></option>
          <%for i=1 to 12%>
          <option value="<%=i%>" <%If izhim=cstr(i) then%>selected<%End If%>><%=i%></option>
          <%next%>
        </select>
        月</td>
      <td>订单号:
        <input name="ddh" size="15" value="<%=strddh%>"></td>
      <td>流水号: <input name="lsh" size="15" value="<%=strlsh%>"></td>
    </tr>
    <tr>
      <td>任务类型:
        <select name="rwlr">
          <option value="0" selected="selected">全部</option>
          <option value="1" <%If strrwlr="1" Then%>selected="selected"<%End If%>>全套</option>
          <option value="2" <%If strrwlr="2" Then%>selected="selected"<%End If%>>机头</option>
          <option value="3" <%If strrwlr="3" Then%>selected="selected"<%End If%>>定型</option>
          <option value="4" <%If strrwlr="4" Then%>selected="selected"<%End If%>>新品</option>
          <option value="5" <%If strrwlr="5" Then%>selected="selected"<%End If%>>非调</option>
          <option value="6" <%If strrwlr="6" Then%>selected="selected"<%End If%>>TB</option>
        </select>
        &nbsp;&nbsp;任务状态:
        <select name="rwzt">
          <option value="0" selected="selected">全部</option>
          <option value="1" <%If strrwzt="1" Then%>selected="selected"<%End If%>>正在调试</option>
        <%If StrRwlx=1 Then%>
          <option value="2" <%If strrwzt="2" Then%>selected="selected"<%End If%>>调试结束</option>
          <option value="3" <%If strrwzt="3" Then%>selected="selected"<%End If%>>入库打包</option>
          <option value="4" <%If strrwzt="4" Then%>selected="selected"<%End If%>>出库发货</option>
         <%else%>
          <option value="2" <%If strrwzt="2" Then%>selected="selected"<%End If%>>验收合格</option>
         <%End If%>
        </select></td>
      <td>客户名:
        <input name="khm" size="15" value="<%=strkhm%>">
        <input type=hidden name=rwlx id=rwlx value=<%=StrRwlx%>></td>
        <td align="center"><input class=btn_2k3 type="submit" value="检 索" /></td>
    </tr>
  </form>
</table>
<%
End Function
%>
