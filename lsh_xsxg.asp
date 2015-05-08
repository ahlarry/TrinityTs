<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<SCRIPT language=Javascript src="Inc/jsfunc.js"></SCRIPT>
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

'定义搜索参数
Dim strlsh, strddh, strkhm, strcz, StrRwlx, strFeedBack
strlsh=Trim(Request("lsh"))
If instr(strlsh,"，")>0 Then 	'如包含有全角符号转换为半角
	strlsh=Replace(strlsh,"，",",")
End If
strddh=Trim(Request("ddh"))
strkhm=Trim(Request("khm"))
strcz=Trim(Request("cz"))
StrRwlx=Trim(Request("rwlx"))
strFeedBack="lsh="&strlsh&"&ddh="&strddh&"&khm="&strkhm
If StrRwlx="" Then StrRwlx="1" End If
%>
<style type="text/css">
<!--
.STYLE1 {
	color: #EEEEEE;
}
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
  <tbody>
    <tr>
      <td valign="bottom" background="imgV2/left_bg1.png" class="bg_x">&nbsp;</td>
      <td width="9" align="right" valign="top" background="imgV2/left_bg2.png"><img src="imgV2/left_yy.png" width="9" height="510" /></td>
      <td width="211" valign="top" background="imgV2/left_bg.gif"><!--#include file="login.asp" -->
      <table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td class="title_left" height="34" align="center" ><font color="#000000">分 值 统 计</font></td>
        </tr>
        <tr>
          <td align="center"><table width="86%" border="0" align="center" cellpadding="0" cellspacing="2">
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='miss_acc.asp'>任 务 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='assess_dis.asp'>考 评 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='acc_list.asp'>厂 内 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='extacc_list.asp'>厂 外 分 值</a></TD>
              </tr>
              <tr>
                <TD width=34 height=24 align=center><IMG src="img/class1.gif" width=20 height=20></TD>
                <TD width="100%"><a href='lsh_xsxg.asp'>重 置 系 数</a></TD>
              </tr>
            </table></td>
        </tr>
      </table>

        <!--#include file="Left_pld.asp" --></td>
      <td width="741" align="center" valign="top" background="Images_v1/rith_bg.gif" bgcolor="#F8FFE8"><table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tbody>
            <tr>
              <td width="21"><img src="Images_v1/yx2_1.gif" width="35" height="21" /></td>
              <td background="Images_v1/yx2_2.gif">&nbsp;</td>
              <td width="21" align="right"><img src="Images_v1/yx2.png" width="247" height="21" /></td>
            </tr>
          </tbody>
        </table>
        <table width="98%" border="0" cellspacing="0" cellpadding="0">
          <tbody>
            <tr>
              <td width="40%" height="18">　计划任务&gt;&gt; 更改任务系数 </td>
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
        <table width="98%" border="2" bordercolor="#33FFCC">
          <%If Strlsh="" Then
			Call Mlist()
		else
			Call MAssign()%>
          <%
		End If %>
        </table></td>
    </tr>
  </tbody>
</table>
<!-- #include file="Inc/Foot.asp" -->
<%
Function SearchInfo()
%>
<table border="0" cellpadding="2" cellspacing="0" width="100%">
  <form action="<%=Request.Servervariables("SCRIPT_NAME")%>" method="get" name="frm_searchinfo" id="frm_searchinfo" onsubmit='return true;'>
    <tr>
      <td>&nbsp;&nbsp;输入筛选条件:</td>
      <td>订单号:
        <input name="ddh" value="<%=strddh%>"></td>
      <td>流水号:
        <input name="lsh" size="10" value="<%=strlsh%>"></td>
      <td>客户名:
        <input name="khm" size="10" value="<%=strkhm%>"></td>
      <td><input type="hidden" name="rwlx" value="<%=strrwlx%>">
        <input class=btn_2k3 type="submit" value="检 索" /></td>
    </tr>
  </form>
</table>
<%
End Function

Function Mlist()
strSql=""
	If strLsh <> "" Then
		strSql = "and lsh like '%"&strLsh&"%' "
	End If
	If strddh <> "" Then
		strSql = "and  ddh like '%"&strddh&"%' " & strSql
	End If
	If strkhm <> "" Then
		strSql = "and  khmc like '%"&strkhm&"%' " & strSql
	End If

set rs=server.createobject("adodb.recordset")
	sqltex="select * from mission where datediff('m',cksj,now())<4 "&strsql&" Order by lsh desc,jhjssj"
rs.open sqltex,conn,1,1
dim PerPage
If Strlsh="" Then
	PerPage=20
else
	PerPage=10
End If
'假如没有数据时
If rs.eof and rs.bof then
response.write "<p align='center'><font color='#ff0000'>没有找到相关任务！</font></p>"
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
<table width="98%" border="1" cellpadding="0" cellspacing="0"  bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2">已完成待修改任务列表</span><br />
  </caption>
  <tbody>
    <tr>
      <th width="3%" class="STYLE4">ID</th>
      <th class="STYLE4">订单号</th>
      <th class="STYLE4">流水号</th>
      <th class="STYLE4">客户名称</th>
      <th class="STYLE4">任务类型</th>
      <th class="STYLE4">额定用料</th>
      <th class="STYLE4">实际用料</th>
      <th class="STYLE4">调试次数</th>
      <th class="STYLE4">调试结束</th>
      <th class="STYLE4">出库时间</th>
      <th class="STYLE4">操作</th>
    </tr>
    <%if not rs.eof then
i=1
do while not rs.eof
%>
  <form name="frm_searchlsh" action='<%=Request.ServerVariables("SCRIPT_NAME") & "?rwlx=" &Strrwlx& "&Lsh=" & Rs("lsh")%>' method="post" onSubmit="return searchlsh_true();">
    <tr>
      <td class="STYLE4"><div align="center"> <%=i%></div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("ddh")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <a href='intpld_dis.asp?s_lsh=<%=rs("lsh")%>'><%=Rs("lsh")%>&nbsp;</a></div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("khmc")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("rwlx")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("edtsyl")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjtsyl")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjtscs")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjjssj")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("cksj")%>&nbsp;</div></td>
      <td class="STYLE4"><label>
        <div align="center">
          <input type="submit" name="Submit2" value="修改" />
          <input type="hidden" name="AssLsh" value='<%=Rs("lsh")%>' />
        </div>
        </label>
        <div align="center"> </div></td>
    </tr>
  </form>
  <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
  </tbody>

</table>
<%
      Call showpage("lsh_xsxg.asp?"&strFeedBack,Rs.RecordCount,PerPage,True,True,"条")
      %>
<%End If
  End Function

  Function MAssign()
	Dim m, iscxs, itsxs
	'列数，修改后生产系数，修改后调试系数
	iscxs=0 : itsxs=0
	strSql="select * from [ins_tsfz] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If Not(Rs.eof Or Rs.bof) Then
%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <tr>
    <th scope="col">ID</th>
    <th scope="col">流水号</th>
    <th scope="col">调试员</th>
    <th scope="col">任务内容</th>
    <th scope="col">分值</th>
    <th scope="col">备注</th>
  </tr>
  <%For m=1 to Rs.RecordCount%>
  <tr>
    <td><div align="center"><%=m%></div></td>
    <td><div align="center"><%=Rs("lsh")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("zrr")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("js")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("fz")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("bz")%>&nbsp;</div></td>
  </tr>
  <%
  if Rs("js")="生产" and not(isnull(Rs("bz"))) Then iscxs=Mid(Rs("bz"),Instr(Rs("bz"),"→")+1,Instr(Rs("bz"),")")-Instr(Rs("bz"),"→")-1) End If
  if Rs("js")="调试" and not(isnull(Rs("bz"))) Then itsxs=Mid(Rs("bz"),Instr(Rs("bz"),"→")+1,Instr(Rs("bz"),")")-Instr(Rs("bz"),"→")-1) End If

  Rs.movenext
  next%>
</table>
<br>
<%End If
    Rs.Close
	Dim strEdcs, strSjcs, strCsxs, strYlxs, Tmpxs
	'根据用料、调试次数、和出厂方式求的总的奖励
	strSql="select * from [mission] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If not(Rs.eof Or Rs.bof) Then
		strEdcs=Split(Rs("edtscs"),":")
		strSjcs=Split(Rs("sjtscs"),":")
		strCsxs=(strEdcs(0)-strSjcs(0))*0.02+(strEdcs(1)-strSjcs(1))*0.05	'调试次数系数
		strYlxs=Fix((Rs("edtsyl")-Rs("sjtsyl"))/(Rs("edtsyl")*0.25))*0.05	'用料系数
		Select Case Rs("ccfs")
			Case "产品合格"
				Tmpxs=1+strCsxs+strYlxs
			Case "一项不合格"
				Tmpxs=1+strCsxs+strYlxs-0.1
			Case "两项不合格"
				Tmpxs=1+strCsxs+strYlxs-0.2
			Case "三项不合格"
				Tmpxs=1+strCsxs+strYlxs-0.3
			Case "通知放行"
				Tmpxs=0.6
			Case "粗调"
				Tmpxs=0.6
		End Select
	End If
	Rs.Close
	If Tmpxs<0.6 Then Tmpxs=0.6 End If
	If iscxs=0 Then iscxs=Tmpxs End If
	If itsxs=0 Then itsxs=Tmpxs End If
%>
<form name="Form_ExtF" method="post" action="xsxg_indb.asp"  onsubmit='return true;'>
  <table id="Ext_FzTb" width='90%' border="2" bordercolor="#33FFCC">
    <caption>
    <span class="STYLE2">修改<%=strlsh%> 厂内任务系数</span><br />
    </caption>
    <tr>
    	<th scope="col">任务</th><th scope="col">默认系数</th><th scope="col">修改系数</th>
    </tr>
    <tr>
    	<td align="center" >生产</td><td align="center" ><%=Tmpxs%></td>
    	<td align="center" ><input name="scxs" value="<%=iscxs%>" onKeyPress="javascript:validationNumber(this, 'u_float', 2, '');"></td>
    </tr>
    <tr>
      <td align="center" >调试</td><td align="center" ><%=Tmpxs%></td>
      <td align="center" ><input name="tsxs" value="<%=itsxs%>" onKeyPress="javascript:validationNumber(this, 'u_float', 2, '');"></td>
    </tr>
    <tr>
      <td align="center" colspan="3"><input type="hidden" id="yscxs" name="yscxs" value="<%=iscxs%>"/>
      <input type="hidden" id="ytsxs" name="ytsxs" value="<%=itsxs%>"/>
      <input type="hidden" id="ilsh" name="ilsh" value="<%=strlsh%>"/>
        <input type="Submit" name="Submit6" value="∷ 确 定 ∷" /></td>
    </tr>
  </table>
</form>
<%
  End Function
%>