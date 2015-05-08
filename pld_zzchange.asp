<!--#include file="Inc/Sysconn.asp" -->
<!--#include file="inc/Skin_CSS.asp"-->
<!--#include file="Inc/user_dbinf.asp"-->
<%
If session("UserName")="" then
	response.redirect "UserServer.asp"
End If

Dim strFeedBack,StrRwlx,StrLsh
StrRwlx=Trim(Request("rwlx"))
StrLsh=Trim(Request("lsh"))
If StrRwlx="" Then StrRwlx="1" End If
strFeedBack="&rwlx="&StrRwlx&"&s_lsh="&StrLsh
%>
<style type="text/css">
<!--
.STYLE1 {
	color: #EEEEEE
}
.STYLE2 {
	font-size: x-large;
	font-weight: bold;
	color: #000066;
}
.STYLE4 {
	color: #000066;
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
        <tr>
          <td width="29%" height="18">　计划任务&gt;&gt;更改任务责任人</td>
          <td width="71%"><div align="right">
              <label>
              <input name="rwlx" type="radio" value="1" onClick="location.href='pld_zzchange.asp?rwlx='+ this.value" <%If strrwlx="1" Then%>="" checked="checked" <%End If%>="" />
              厂内调试
              <input name="rwlx" type="radio" value="0" onClick="location.href='pld_zzchange.asp?rwlx='+ this.value" <%If strrwlx="0" Then%>="" checked="checked" <%End If%>="" />
              厂外调试</label>
            </div></td>
        </tr>
        <tr>
          <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </table>
      <Table border=0 cellpadding=2 cellspacing=0 width="98%">
        <Form name=frm_searchlsh action='<%=Request.ServerVariables("SCRIPT_NAME") & "?rwlx=" &Strrwlx%>' method=post onsubmit='return searchlsh_true();'>
          <tr>
            <td>&nbsp;&nbsp;输入流水号:
              <input tabindex=1 type=text name=lsh size=15 value="<%=Trim(Request("lsh"))%>">
              <input type="submit" value=" 查 找 ">
            </td>
          </tr>
        </Form>
      </Table>
      <Table border=0 cellpadding=2 cellspacing=0 width="98%">
        <tr>
          <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </Table>
      <%Call mtask_change(StrRwlx,StrLsh)%>
      <Table border=0 cellpadding=2 cellspacing=0 width="98%">
        <tr>
          <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
        </tr>
        <tr>
          <td height="1" colspan="2">&nbsp;</td>
        </tr>
      </Table>
      <%Call mtask_list()%>
    </td>
  </tr>
</table>
<!-- #include file="Inc/Foot.asp" -->
<map name="Map">
  <area shape="rect" coords="18,14,69,24" href="News.asp">
</map>
</BODY>
</HTML><%
Function mtask_change(x,y)
Dim m
If y="" Then Responsr.Write("请输入要更改的任务书的流水号!") : Exit Function
If x="1" Then
strSql="select * from [mission] where lsh='"&y&"' and isNull(rksj)"
else
strSql="select * from [mission] where lsh='"&y&"' and isNull(cwjssj)"
End If
set Rs=server.createobject("adodb.recordset")
Rs.open strSql,Conn,1,1
If Rs.Eof Or Rs.Bof Then
errmsg="流水号为 【 " & y & " 】 的任务书不存在或已完成!"
Call WriteErrMsg()
Exit Function
End If
Rs.Close
%>
<form name="Form_Assing" id="Form_Assing" method="post" action="pld_assindb.asp?action=Change">
  <table width="98%" border="0">
  <tr>
  <td>
  <fieldset>
  <legend>已分配</legend>
  <table width="98%" border="0">
    <tr>
      <td class="STYLE4">流水号:</td>
      <td class="STYLE4"><font color="red"><%=y%></font></td>
    </tr>
    <%	If x="1" Then
strSql="select * from [ins_tsxx] where lsh='"&y&"'"
else
strSql="select * from [ext_tsxx] where lsh='"&y&"'"
End If
set Rs=server.createobject("adodb.recordset")
Rs.open strSql,Conn,1,1
m=Rs.recordcount
For i=1 to m
%>
    <tr>
      <td class="STYLE4">责任人:</td>
      <td><span id="tsxr<%=i%>" class="STYLE4"><%=Rs("tsy")%></span>
        <input type="hidden" name="zrr<%=i%>" value=<%=Replace(Rs("tsy"),", ",",")%> /></td>
      <input type="hidden" name="id<%=i%>" value=<%=Rs("id")%> />
    </td>
    </tr>
    <%Rs.movenext 
Next
Rs.Close%>
    <tr>
      <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
    </tr>
    <tr>
      <input type="hidden" name="num" value="<%=m%>" />
    </tr>
    <tr>
      <td colspan="2" align="center"><input name="submit" type="submit" value=" · 确 定 · " />
    </tr>
  </table>
</form>
</fieldset>
</td>
<td><fieldset>
  <legend>调整</legend>
  <table width="98%" border="0">
    <tr>
      <td rowspan="3"><div align="center">
          <select name="dxtsy" style="height:120px;width:100px" multiple="multiple" id="dxtsy">
            <%for i = 0 to ubound(c_allzy)%>
            <option value="<%=c_allzy(i)%>"><%=c_allzy(i)%></option>
            <%next%>
          </select>
        </div></td>
      <td><div align="center">第
          <select name="cishu" id="cishu">
            <%for i = 1 to m%>
            <option value="<%=i%>"><%=i%></option>
            <%next%>
          </select>
          次</div></td>
    </tr>
    <tr>
      <td><div align="center">
          <input name="button2" type="button" id="button2" onClick="ChanTsy('del');" value="删除"/>
        </div></td>
    </tr>
    <tr>
      <td><div align="center">
          <input name="button" type="button" id="button" onClick="ChanTsy('cha');" value="修改"/>
        </div></td>
    </tr>
  </table>
  </fieldset></td>
</tr>
</table>
<%
End Function

Function mtask_list		'未完成任务列表
set rs=server.createobject("adodb.recordset")
If StrRwlx="1" Then
	sqltex="select * from [mission] where isNull(rksj) Order BY jhjssj desc"
else
	sqltex="select * from [mission] where isNull(cwjssj)"
End If
rs.open sqltex,conn,1,1
dim PerPage
PerPage=20
'假如没有数据时
If rs.eof and rs.bof then
response.write "<p align='center'><font color='#ff0000'>没有待分配任务！</font></p>"
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
<table width="98%" border="2" bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2">厂内正在调试任务列表</span><br />
  </caption>
  <tr>
    <th width="3%" class="STYLE4">ID</th>
    <th width="17%" class="STYLE4">订单号</th>
    <th width="10%" class="STYLE4">流水号</th>
    <th class="STYLE4">客户名称</th>
    <th width="11%" class="STYLE4">调试次数</th>
    <th width="15%" class="STYLE4">调试结束</th>
    <th width="15%" class="STYLE4">入库时间</th>
    <th width="12%" class="STYLE4">操作</th>
  </tr>
  <%if not rs.eof then
i=1
do while not rs.eof

%>
  <form name="frm_changezrr" action='<%=Request.ServerVariables("SCRIPT_NAME") & "?rwlx=" &Strrwlx& "&lsh=" & Rs("lsh")%>' method="post" onSubmit="return searchlsh_true();">
    <tr>
      <td class="STYLE4"><div align="center"><%=i%></div></td>
      <td class="STYLE4"><div align="center"><%=Rs("ddh")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"><a href="intpld_dis.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("lsh")%></a></div></td>
      <td class="STYLE4"><div align="center"><%=Rs("khmc")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center"><%=Rs("sjtscs")%>&nbsp;</div></td>
      <td class="STYLE4"><div align="center">&nbsp;</div></td>
      <td class="STYLE4"><div align="center">&nbsp;</div></td>
      <td class="STYLE4"><label>
        <div align="center">
          <input type="submit" name="Submit2" value="更改责任人" />
          </label>
        </div></td>
    </tr>
  </form>
  <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
</table>
<%else%>
<table width="98%" border="2" bordercolor="#33FFCC">
  <caption>
  <span class="STYLE2">厂外正在调试任务列表</span><br />
  </caption>
  <tr>
    <th scope="col">ID</th>
    <th scope="col">流水号</th>
    <th scope="col">客户名</th>
    <th scope="col">计划结束</th>
    <th scope="col">实际结束</th>
    <th scope="col">验收结果</th>
    <th scope="col">操作</th>
  </tr>
  <%if not rs.eof then
i=1
do while not rs.eof
%>
  <form name="frm_changezrr" action='<%=Request.ServerVariables("SCRIPT_NAME") & "?rwlx=" &Strrwlx& "&lsh=" & Rs("lsh")%>' method="post" onSubmit="return searchlsh_true();">
    <tr>
      <td><div align="center"><%=i%></div></td>
      <td><div align="center"><a href="extpld_dis.asp?s_lsh=<%=rs("lsh")%>"><%=Rs("lsh")%>&nbsp;</a></div></td>
      <td><div align="center"><%=Rs("khmc")%>&nbsp;</div></td>
      <td><div align="center"><%=Rs("jhjssj")%>&nbsp;</div></td>
      <td><div align="center"><%=Rs("sjjssj")%>&nbsp;</div></td>
      <td><div align="center"><%=Rs("ysjg")%>&nbsp;</div></td>
      <td><div align="center">
          <input type="submit" name="Submit2" value="更改责任人" />
        </div></td>
    </tr>
  </form>
  <%
i=i+1
if i > Perpage then exit do
rs.movenext
loop
end if
%>
</table>
<%End If%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr>
    <td><div align="right">
        <%
Response.write "全部"
Response.write "共" & Cstr(Rs.RecordCount) & "个任务&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
Response.write "第" & Cstr(CurrentPage) &  "/" & Cstr(rs.pagecount) & "&nbsp;"
If currentpage > 1 Then
response.write "<a href='pld_list.asp?&page="+cstr(1)+"'>&nbsp;首页&nbsp;</a>"
Response.write "<a href='pld_list.asp?page="+Cstr(currentpage-1)+"'>&nbsp;上一页&nbsp;</a>"
Else
Response.write "&nbsp;上一页&nbsp;"
End if
If currentpage < Rs.PageCount Then
Response.write "<a href='pld_list.asp?page="+Cstr(currentPage+1)+"'>&nbsp;下一页&nbsp;</a>"
Response.write "<a href='pld_list.asp?page="+Cstr(Rs.PageCount)+"'>&nbsp;尾页&nbsp;</a>"
Else
Response.write ""
Response.write "&nbsp;下一页&nbsp;"
End if
Response.write "转到第"
response.write"<select name='sel_page' onChange='javascript:location=this.options[this.selectedIndex].value;'>"
    for i = 1 to Rs.PageCount
       if i = currentpage then 
          response.write"<option value='pld_list.asp?page="&i&"&id="&id&"' selected>"&i&"</option>"
       else
          response.write"<option value='pld_list.asp?page="&i&"&id="&id&"'>"&i&"</option>"
       end if
    next
response.write"</select>页"
rs.close
%>
      </div></td>
  </tr>
  <tr>
    <td bgcolor="#00CC00"></td>
  </tr>
</table>
<%End If%>
<%End Function%>
<script language="javascript">
var i; 
function ChanTsy(arg)
{
	var strcs=document.all.cishu.value;
	var strtsy=eval("document.all.zrr"+ strcs +".value;");
	if (strtsy != ""){
		strtsy=strtsy.split(",");
	}
	else {
		strtsy=new Array();
	}
	
    for (i=0;i<document.all.dxtsy.length;i++){
        if (document.all.dxtsy.options[i].selected == true){
        	strtsy.push(document.all.dxtsy.options[i].text);
        } 
    }
    for(i = 0; i < strtsy.length; i++ ){
		for( var j = strtsy.length - 1; j > i; j-- ){
			if( strtsy[j] == strtsy[i] ){
				strtsy.splice(j,1);
				strtsy.splice(i,1);
			}
		}
	}
	if (arg == "cha"){
		eval("document.all.tsxr"+ strcs +".innerHTML = strtsy");
		eval("document.all.zrr"+ strcs +".value = strtsy");
	}
	else {
		eval("document.all.tsxr"+ strcs +".innerHTML = ''");
		eval("document.all.zrr"+ strcs +".value = ''");
}
}
</script>
