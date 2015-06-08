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
strFeedBack="rwlx="&StrRwlx&"&lsh="&strlsh&"&ddh="&strddh&"&khm="&strkhm&"&cz="&strcz
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
              <td width="40%" height="18">　计划任务&gt;&gt; 分配调试任务 </td>
              <td width="71%"><div align="right">
                  <label>
                  <input name="rwlx" type="radio" value="1" onClick="location.href='pld_assign.asp?rwlx='+ this.value" <%If strrwlx="1" Then%>="" checked="checked" <%End If%>="" />
                  厂内调试
                  <input name="rwlx" type="radio" value="0" onClick="location.href='pld_assign.asp?rwlx='+ this.value" <%If strrwlx="0" Then%>="" checked="checked" <%End If%>="" />
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
        <table width="98%" border="2" bordercolor="#33FFCC">
          <%If Strlsh="" Then
			Call Mlist()
		else
			Call MAssign()%>
          <table width="98%" border="0" cellspacing="0" cellpadding="0">
            <tbody>
              <tr>
                <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
              </tr>
              <tr>
                <td height="1" colspan="2">&nbsp;</td>
              </tr>
            </tbody>
          </table>
          <%Call Mlist()
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
      <td>操作内容:
        <select name="cz">
          <option value="" selected="selected">全部</option>
          <option value="0" <%If strcz="0" Then%>selected="selected"<%End If%>>调试</option>
          <option value="1" <%If strcz="1" Then%>selected="selected"<%End If%>>用料</option>
          <option value="2" <%If strcz="2" Then%>selected="selected"<%End If%>>入库</option>
          <option value="3" <%If strcz="3" Then%>selected="selected"<%End If%>>出库</option>
        </select></td>
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
	Select case strcz
		case "0"
			strSql = "and isNull(sjjssj) and rwlx<>'非调' and rwlx<>'TB' " & strSql
		case "1"
			strSql = "and not(isNull(sjjssj)) and sjtsyl=0 " & strSql
		case "2"
			strSql = "and isNull(rksj) and (sjtsyl<>0 or rwlx='非调' or rwlx='TB') " & strSql
		case "3"
			strSql = "and not(isNull(rksj)) and isNull(cksj)" & strSql
	End Select

set rs=server.createobject("adodb.recordset")
If StrRwlx="0" Then
	sqltex="select * from mission where not(isNull(cksj)) and isNull(cwjssj) Order by lsh"
else
	sqltex="select * from mission where isNull(cksj) "&strsql&" Order by jhjssj,lsh"
End If
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
  <span class="STYLE2">未分配调试任务列表</span><br />
  </caption>
  <tbody>
    <tr>
      <th width="3%" class="STYLE4">ID</th>
      <th class="STYLE4">订单号</th>
      <th class="STYLE4">流水号</th>
      <th class="STYLE4">客户名称</th>
      <%If StrRwlx="1" Then %>
      <th class="STYLE4">任务类型</th>
      <th class="STYLE4">额定用料</th>
      <th class="STYLE4">实际用料</th>
      <th class="STYLE4">调试次数</th>
      <th class="STYLE4">调试结束</th>
      <th class="STYLE4">入库时间</th>
      <%else%>
      <th class="STYLE4">出库时间</th>
      <th class="STYLE4">开始调试</th>
      <th width="30%" class="STYLE4">调试信息</th>
      <%End If%>
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
      <%If StrRwlx="1" Then %>
      <td class="STYLE4"><div align="center"> <%=Rs("rwlx")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("edtsyl")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjtsyl")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjtscs")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("sjjssj")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("rksj")%>&nbsp;</div></td>
      <%else%>
      <td class="STYLE4"><div align="center"> <%=Rs("cksj")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"> <%=Rs("cwkssj")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center" alt="<%=Rs("bz")%>"> <%=cutstr(Rs("bz"),10)%>&nbsp;</div></td>
      <%End If%>
      <td class="STYLE4"><label>
        <div align="center">
          <%if StrRwlx="1" Then%>
          <input type="submit" name="Submit2" <%If isNull(Rs("sjjssj")) Then%>value="分配"<%elseif isNull(Rs("rksj")) Then%>value="入库"<%else%>value="出库"<%End If%> />
          <%else%>
          <input type="submit" name="Submit2" value="分配" />
          <%End If%>
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
      Call showpage("pld_assign.asp?"&strFeedBack,Rs.RecordCount,PerPage,True,True,"条")
      %>
<%End If
  End Function

  Function MAssign()
  %>
<%
  If StrRwlx="1" Then
		Call Assed_ins()
		strSql="select * from [mission] where lsh='"&strlsh&"' and isNull(cksj)"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If Not(Rs.eof Or Rs.bof) Then
			If IsNull(Rs("sjjssj")) and Rs("rwlx")<>"非调" and Rs("rwlx")<>"TB" Then
				Call InsAssign()
			else
				if Rs("sjtsyl")=0  and Rs("rwlx")<>"非调" and Rs("rwlx")<>"TB" Then
					Call YlAssign()
				else
					if isnull(Rs("rksj")) Then
						Call RkAssign()
					else
						If IsNull(Rs("cksj")) Then
						 	Call CkAssign()
						End If
					End If
				End If
			End if
		else
			Response.write("<font color='red'>任务不存在或已完成！</font>")
		End If
		Rs.Close
	else
		Dim TmpArr,strjdu
		TmpArr=split(strlsh,",")	'多个序列号同时分配
		strjdu="" : errmsg=""
		For i=0 to ubound(TmpArr)
			Call Assed_ext(TmpArr(i))
			strSql="select * from [mission] where lsh='"&TmpArr(i)&"' and isNull(cwjssj)"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			If Rs.eof Or Rs.bof Then
				errmsg=errmsg&"流水号 <strong>"&TmpArr(i)&"</strong> 任务不存在或已完成!<br>"
			else
				If strjdu="" Then
					strjdu=Rs("ysjg")
				else
					strjdu=strjdu&","&Rs("ysjg")
				End If
			End If
		Next
		Dim Tmpjdu,j
'		errmsg=errmsg&strjdu
		strjdu=split(strjdu,",")
		For i=0 to ubound(strjdu)
			For j=1 to ubound(strjdu)
				if strjdu(i)<>strjdu(j) then
					errmsg=errmsg&"任务进度不一致，无法合成一批录入!<br>"
					exit for
				End If
			Next
			if instr(errmsg,"进度不一致")>0 then exit for
		Next
		If errmsg<>"" Then
			Call WriteErrMsg()
		else
			If strjdu(0)="正在调试" Then
				Call ExtAssign(strlsh)
			else
				Call ExtFzindb(strlsh)
			End If
		End If
	End If
End Function

Function Assed_ins()	'已分配厂内任务
	Dim m
	strSql="select * from [ins_tsxx] where lsh='"&strlsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If Not(Rs.eof Or Rs.bof) Then
%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <tr>
    <th scope="col">ID</th>
    <th scope="col">调试员</th>
    <th scope="col">调试内容</th>
    <th scope="col">出厂方式</th>
  </tr>
  <%For m=1 to Rs.RecordCount%>
  <tr>
    <td><div align="center"><%=m%></div></td>
    <td><div align="center"><%=Rs("tsy")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("tsnr")%>&nbsp;</div></td>
    <td><div align="center"><%=Rs("cc")%>&nbsp;</div></td>
  </tr>
  <%Rs.movenext
  next%>
</table>
<br>
<%End If
    Rs.Close
  End Function

Function Assed_ext(ilsh)	'已分配厂外任务
	strSql="select * from [ext_tsxx] where lsh='"&ilsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	If Not(Rs.eof Or Rs.bof) Then
%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <tr>
    <th scope="col">姓名</th>
    <th scope="col">开始时间</th>
    <th scope="col">结束时间</th>
    <th scope="col">验收</th>
    <th scope="col">姓名</th>
    <th scope="col">开始时间</th>
    <th scope="col">结束时间</th>
    <th scope="col">验收</th>
  </tr>
  <%do while Not(Rs.eof Or Rs.bof)%>
  <tr>
    <td><div align="center"><%=Rs("tsy")%></div></td>
    <td><div align="center"><%=Rs("kssj")%></div></td>
    <td><div align="center"><%=Rs("jssj")%></div></td>
    <td><div align="center"><%=Rs("cc")%></div></td>
    <%Rs.movenext
    If Not(Rs.eof Or Rs.bof) Then%>
    <td><div align="center"><%=Rs("tsy")%></div></td>
    <td><div align="center"><%=Rs("kssj")%></div></td>
    <td><div align="center"><%=Rs("jssj")%></div></td>
    <td><div align="center"><%=Rs("cc")%></div></td>
    <%Rs.movenext
    else%>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <td><div align="center">&nbsp;</div></td>
    <%End If%>
  </tr>
  <%loop%>
</table>
<br>
<%End If
  Rs.Close
  End Function

  Function InsAssign()	'厂内任务分配%>
<form name="Form_Assing" id="Form_Assing" method="post" action="pld_assindb.asp?action=Massign">
  <table width="70%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center"><strong>待选调试员</strong></div></td>
      <td rowspan="2"><div align="center">
          <INPUT language="javascript" name="btn_select_addall" onClick="fun_select_addall(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title="添加所有" type=button value=: DESIGNTIMESP="17713">
          <br>
          <br>
          <INPUT language="Javascript" name="btn_select_addany" onClick="fun_select_addany(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title="添加所选" type=button value="8">
          <br>
          <br>
          <INPUT language="javascript" name="btn_select_dltany" onClick="fun_select_dltany(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title ="删除所选" type=button value="7">
          <br>
          <br>
          <INPUT language="javascript" name="btn_select_dltall" onClick="fun_select_dltall(document.Form_Assing)" style="COLOR: #000066; FONT-FAMILY: Webdings; FONT-SIZE: 12pt; FONT-WEIGHT: normal; HEIGHT: 28px; WIDTH: 27px" title ="删除所有" type=button value="9">
        </div>
        <div align="center"></div></td>
      <td><div align="center"><strong>已选调试员</strong></div></td>
    </tr>
    <tr>
      <td><div align="center">
          <select name="dxtsy" style="height:160px;width:100px" multiple="multiple" id="dxtsy">
            <%for i = 0 to ubound(c_allzy)%>
            <option value="<%=c_allzy(i)%>"><%=c_allzy(i)%></option>
            <%next%>
          </select>
        </div></td>
      <td><div align="center">
          <select name="yxtsy" style="height:160px;width:100px" multiple="multiple" id="yxtsy" >
            <%
          	strSql="select * from [ins_tsxx] where lsh='"&strlsh&"' order by Id desc"
			set Rs=server.createobject("adodb.recordset")
			Rs.open strSql,Conn,1,1
			If Not(Rs.eof Or Rs.bof) Then
				StrTsy=split(Rs("tsy"),", ")
		        for i = 0 to ubound(StrTsy)%>
            <option value="<%=StrTsy(i)%>"><%=StrTsy(i)%></option>
            <%next
			End If
			Rs.Close
          %>
          </select>
        </div></td>
    </tr>
    <tr>
      <td colspan="3"><div align="center">调试内容:
          <select name="tsnr">
            <option value="全套">全套</option>
            <option value="模头">模头</option>
            <option value="定型">定型</option>
          </select>
          &nbsp;&nbsp;参与次数:
          <input type="text" name="cycs" size="5" value="1" onKeyPress="javascript:validationNumber(this, 'u_float', 100, '');">
          &nbsp;&nbsp;验收方式:
          <select name="ccfs">
            <option value="正在调试" selected>正在调试</option>
            <option value="粗调">粗调</option>
            <option value="产品合格">产品合格</option>
            <option value="一项不合格">一项不合格</option>
            <option value="两项不合格">两项不合格</option>
            <option value="三项不合格">三项不合格</option>
            <option value="通知放行">通知放行</option>
          </select>
        </div></td>
    </tr>
  </table>
  <br>
  <label>
  <div align="center">
    <input type="Submit" name="Submit3" onClick="SelTsy(document.Form_Assing);" value="∷确 定∷" />
    <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
    <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
  </div>
  </label>
</form>
<%End Function

Function YlAssign()		'用料
%>
<form name="Form_YL" id="Form_YL" method="post" action="pld_assindb.asp?action=YL">
  <table width="100%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center">试模料用量：
          <input type="text" name="smyl" />
          &nbsp;&nbsp;
          <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='yxtsyjs';
  		myDate.display();
		</script>
          &nbsp;&nbsp;
          <input type="Submit" name="Submit4" value="∷确 定∷" />
          <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
          <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
        </div></td>
    </tr>
  </table>
</form>
<%
End Function

Function RkAssign()		'入库
%>
<form name="Form_Ruku" id="Form_Ruku" method="post" action="pld_assindb.asp?action=Ruku">
  <table width="100%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center">入库时间：
          <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='rksj';
  		myDate.display();
		</script>
          &nbsp;&nbsp;
          <input type="Submit" name="Submit4" value="∷确 定∷" />
          <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
          <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
        </div></td>
    </tr>
  </table>
</form>
<%
End Function

Function CkAssign()		'出库
%>
<form name="Form_Chuuku" id="Form_Chuuku" method="post" action="pld_assindb.asp?action=Chuuku">
  <table width="100%" border="2" bordercolor="#33FFCC">
    <tr>
      <td><div align="center">出库时间：
          <script language=javascript>
  		var myDate=new dateSelector();
  		myDate.year;
 		myDate.inputName='cksj';
  		myDate.display();
		</script>
          &nbsp;&nbsp;
          <input type="Submit" name="Submit4" value="∷确 定∷" />
          <input type="hidden" name="TsLsh" value='<%=strlsh%>' />
          <input type="hidden" name="Tsrwlx" value='<%=strrwlx%>' />
        </div></td>
    </tr>
  </table>
</form>
<%
End Function

Function ExtAssign(lsh)	'厂外任务分配
Dim TmpArr,LshArr,Tmpstr,j,k
TmpArr=split(lsh,",")
For i=0 to ubound(TmpArr)
    strSql="select * from [mission] where lsh='"&TmpArr(i)&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	if not Rs.eof Then
		Select Case Rs("dmdj")
			Case "A类":Tmpstr="1"
			Case "B类":Tmpstr="2"
			Case "C类":Tmpstr="3"
			Case "D类":Tmpstr="4"
			Case "E类":Tmpstr="5"
			Case "F类":Tmpstr="6"
			Case "G类":Tmpstr="7"
			Case "H类":Tmpstr="8"
			Case Else:Tmpstr="9"
		End Select
		TmpArr(i)=TmpArr(i)&"||"&Tmpstr&Left(Rs("mjdj"),1)
	End If
	Rs.Close
Next
%>
<form name="Form_ExtA" id="Form_ExtA" method="post" action="pld_assindb.asp?action=Massign">
  <table width='80%' border="0" cellspacing="0" bordercolor="#33FFCC">
    <tr>
      <td colspan="6" align="center"><span class="STYLE2">分配<%=lsh%>号任务书调试任务</span></td>
    </tr>
    <%
    Tmpstr=ubound(TmpArr)/3
    if int(Tmpstr)<>Tmpstr Then
  		Tmpstr=int(Tmpstr+1)
	end if
	i=0
	For j=0 to Tmpstr
    %>
    <tr>
        <%k=0
       For k=i to i+2
        	If k <= ubound(TmpArr) Then
      			LshArr=split(TmpArr(k),"||")
     			Response.Write("<td align='center'><font color='#FF0000'>"&LshArr(0)&":"&LshArr(1)&"&nbsp;模具系数:<input name='mjxs"&k&"' type='text' size='10'></font></td>")
     		End If
       next
        %>
    </tr>
    <%i=i+3
    next%>
    <tr>
      <td align="left"><font color='#FF0000'>[Ctrl+Enter 新增一行]</font></td>
      <td>&nbsp;</td>
      <td align="right"><font color='#FF0000'>[Shift+Enter 删除一行]</font> </td>
    </tr>
  </table>
  <table id="ExtA_TB" width="80%" border="0" cellspacing="0" bordercolor="#33FFCC">
    <tr id='ExtA_Tr'>
      <td>调试人:
        <select name="zrr" id="zrr">
          <%for i = 0 to ubound(c_allzy)%>
          <option value="<%=c_allzy(i)%>"><%=c_allzy(i)%></option>
          <%next%>
        </select>
      </td>
      <td align="center">开始调试:
        <input name="tskssj" type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10"></td>
      <td align="center"> 结束调试:
        <input name="tsjssj" type="text" onFocus="CalendarWebControl.show(this,false,this.value);" size="10"></td>
      <td><input value="添加" type="button" onClick="Addrow()">
        <input value="删除" type="button" onClick="Delrow()">
      </td>
    </tr>
  </table>
  <table width="70%" border="1" cellspacing="0" bordercolor="#33FFCC">
    <tr>
      <td><span class="STYLE4">备注:</span>
        <input type="hidden" id="TsNum" name="TsNum"/></td>
      <td><textarea name="beiz" cols="65" rows="5"></textarea></td>
    </tr>
    <tr>
      <td colspan="2" align="center">验收：
        <select name="ccfs">
          <option value="正在调试" selected>正在调试</option>
          <option value="验收合格">验收合格</option>
          <option value="通知放行">通知放行</option>
          <option value="拒绝验收">拒绝验收</option>
        </select>
        &nbsp;&nbsp;
        <input type="Submit" name="Submit5" value="∷确 定∷" />
        <input type="hidden" name="TsLsh" value='<%=lsh%>' /></td>
    </tr>
  </table>
</form>
<%
End Function

Function ExtFzindb(slsh)	'厂外任务分值计算(主任打系统)
%>
<form name="Form_ExtF" method="post" action="pld_assindb.asp?action=ExtF" onSubmit='return CheckXs(<%=i%>);'>
  <table id="Ext_FzTb" width='90%' border="0" cellspacing="0" bordercolor="#33FFCC">
    <caption>
    <span class="STYLE2">分配<%=slsh%> 调节系数</span><br />
    </caption>
    <%	'-----------------------------------求厂外调试分值
    Dim strtsfz,ilsh,iid,iyfz,ixs,ifz,ispan
    strtsfz=0 : ilsh=slsh : strlsh=split(slsh,",")
    For i=0 to ubound(strlsh)
		strSql="select * from [mission] where lsh='"&strlsh(i)&"'"
		set Rs=server.createobject("adodb.recordset")
		Rs.open strSql,Conn,1,1
		If Not(Rs.eof Or Rs.bof) Then
			strtsfz=strtsfz + Rs("cwtsfz")*Rs("djxs")		'厂外调试分值
		End If
		Rs.Close
	next
    strSql="select * from [ext_tsfz] where lsh='"&ilsh&"'"
	set Rs=server.createobject("adodb.recordset")
	Rs.open strSql,Conn,1,1
	i=0
	Do while not Rs.eof
	iid="id"&i : iyfz="yfz"&i : ixs="xs"&i : ifz="fz"&i : ispan="span_fz"&i
	%>
    <tr>
      <td>调试员:<%=Rs("zrr")%>
        <input type="hidden" name=<%=iid%> id=<%=iid%> value=<%=Rs("id")%> /></td>
      <td>分值:
        <input type="text" name=<%=iyfz%> id=<%=iyfz%> size="8" value=<%=Rs("fz")%> readonly></td>
      <td>系数:
        <input type="text" name=<%=ixs%> id=<%=ixs%> size="8" onKeyDown="valNum(event);" onpaste="valClip(event);" onChange="ChangeXs(this.value,<%=strtsfz%>,<%=i%>);">
        &nbsp;X&nbsp;<%=strtsfz%></td>
      <td><span id=<%=ispan%>>最终得分:<%=Rs("fz")%></span>
        <input type="hidden" id=<%=ifz%> name=<%=ifz%> value=<%=Rs("fz")%> /></td>
    </tr>
    <%
    Rs.movenext
	i=i+1
	Loop
	Rs.Close
%>
    <tr>
      <td height="1" colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="4" align="center"><input type="hidden" id="ExtFNB" name="ExtFNB" value=<%=i-1%> />
        <input type="hidden" name="TsLsh" value='<%=ilsh%>' />
        <input type="Submit" name="Submit6" value="∷确 定∷" /></td>
    </tr>
  </table>
</form>
<%End Function%>
<script language="javascript">
//厂内调试员选择
function fun_select_addany(theform){
    var i;
    for (i=0;i<theform.dxtsy.length;i++){
        if (theform.dxtsy.options[i].selected == true){
           theform.yxtsy.options[theform.yxtsy.length]=new Option(theform.dxtsy.options[i].text);
        }
    }
    sort_select(theform);
}

function fun_select_addall(theform){
    var i;
    for (i=0;i<theform.dxtsy.length;i++){
        theform.yxtsy.options[theform.yxtsy.length]=new Option(theform.dxtsy.options[i].text);
    }
    sort_select(theform);
}

function fun_select_dltany(theform){
   var i;
   for (i=0;i<theform.yxtsy.length;i++){
       if (theform.yxtsy.options[i].selected == true){
          theform.yxtsy.options[i] =new Option("");
          theform.yxtsy.options.remove(i);
          i--;
       }
   }
   sort_select(theform);
}

function fun_select_dltall(theform){
    var i;
    for (i=0;i<theform.yxtsy.length;i++){
    }
    theform.yxtsy.length=0;
    sort_select(theform);
}

function sort_select(theform){
    var i;
    var right_array = new Array();
    var o = new Object();
    for (i=0;i<theform.yxtsy.length;i++){
        right_array[i] = new Array(theform.yxtsy.options[i].text);
    }
	for (var i=0,j=0; i<right_array.length; i++)
	{
		if (typeof o[right_array[i]] == 'undefined')
		{
			o[right_array[i]] = j++;
		}
	}
	right_array.length = 0;
	for (var key in o)
	{
	right_array[o[key]] = key;
	}
    theform.yxtsy.length=0;
    for (i=0;i<right_array.length;i++){
       		theform.yxtsy.options[theform.yxtsy.length]=new Option(right_array[i]);
    }
    right_array = new Array();
}

function SelTsy(theform)
{
    for (i=0;i<theform.yxtsy.length;i++){
        theform.yxtsy.options[i].selected = true;
    }
}
//增加和删除行
i=1;
function Addrow(){
   var newTR = ExtA_Tr.cloneNode(true);
   newTR.id="a"+(++i)
   ExtA_Tr.parentNode.insertAdjacentElement("beforeEnd",newTR);
   RowReset();
   document.Form_ExtA.TsNum.value=ExtA_TB.rows.length;
}

function RowReset(){
   var RowCount=ExtA_TB.rows.length-1
   var ReName=RowCount
   for (var i=0;i<ExtA_TB.rows[RowCount].cells.length;i++){

      var str=ExtA_TB.rows[RowCount].cells[i].innerHTML
      str=(str.replace(/name=[\d]*/i,"name="+ReName));
      ExtA_TB.rows[RowCount].cells[i].innerHTML=str
   }
}

function Delrow(){
	if (ExtA_TB.rows.length >1){
		ExtA_TB.deleteRow(ExtA_TB.rows.length-1);
	}
	document.Form_ExtA.TsNum.value=ExtA_TB.rows.length;
}

function View(tb){
  for (var r=0;r<tb.rows.length;r++){
     for (var c=0;c<tb.rows[r].cells.length;c++){
        alert(tb.rows[r].cells[c].innerHTML)
     }
  }
}

function ShortcutKey(){
	if((event.ctrlKey)&&(window.event.keyCode==13)) //"Ctrl" + "回车"添加一行
   		Addrow()
	if((event.shiftKey)&&(window.event.keyCode==13)) //"Shift" + "回车"删除一行
		Delrow()
}

document.onkeydown=ShortcutKey;
//厂外调试主任调节系数
function ChangeXs(value,tsfz,n){
	var yfz=eval("document.all.yfz"+ n +".value");
	if(!value)(value=0)
	var zzdf=(parseFloat(yfz) + parseFloat(value)*tsfz).toFixed(2)
	eval("document.all.span_fz"+ n +".innerHTML ='最终得分:"+ zzdf +"'");
	eval("document.all.fz"+ n +".value = '"+ zzdf +"'");
}
//调节系数总和必须为0，以确保总分不会改变
function CheckXs(arg){
	var strsum=0
	for (i=0;i<=arg;i++){
		var xs=eval("document.all.xs"+ i +".value");
		if(!xs)(xs=0)
		strsum = strsum + parseFloat(xs);
	}
	if (strsum!=0){
		alert("系数总和必须为0!")
		return false;
	}
	else
	{
		return true;
	}
}
</script>
