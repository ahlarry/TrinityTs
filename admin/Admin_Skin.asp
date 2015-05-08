<%@language=VBScript codepage=936 %>
<%
option explicit
response.buffer=true
%>
<!--#include file="conn.asp"-->
<!--#include file="../inc/config.asp"-->
<!--#include file="Admin.asp"-->
<!--#include file="inc/function.asp"-->
<!--#include file="../inc/ubbcode.asp"-->
<%
dim rs, sql
dim Action,SkinID,FoundErr,ErrMsg
Action=trim(request("Action"))
SkinID=trim(request("SkinID"))
%>
<!-- #include file="Inc/Head.asp" -->
<table width="600" border="0" align="center" cellpadding="2" cellspacing="1" class="table_southidc">
  <tr class="topbg"> 
    <td class="back_southidc" height="30" colspan="2" align="center"><strong>配 
      色 模 板 管 理</strong></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>管理导航：</strong></td>
    <td><a href="Admin_Skin.asp">配色模板管理首页</a> | <a href="Admin_Skin.asp?Action=Add">添加配色模板</a> 
      | <a href="Admin_Skin.asp?Action=Export">配色模板导出</a> | <a href="Admin_Skin.asp?Action=Import">配色模板导入</a></td>
  </tr>
  <tr class="tdbg"> 
    <td width="70" height="30"><strong>模板说明：</strong></td>
    <td>红色说明为站本站已使用了的CSS</td>
  </tr>
</table>
<%
select case Action
	case "Add","Modify"
		call ShowSkinSet()
	case "SaveAdd"
		call SaveAdd()
	case "SaveModify"
		call SaveModify()
	case "Set"
		call SetDefault()
	case "Del"
		call DelSkin()
	case "Export"
		call Export()
	case "DoExport"
		call DoExport()
	case "Import"
		call Import()
	case "Import2"
		call Import2()		
	case "DoImport"
		call DoImport()
	case else
		call main()
end select
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub main()
	sql="select * from Skin"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,1
%>
<form name="form1" method="post" action="Admin_Skin.asp">
  <table width="600" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="border">
    <tr bgcolor="#FFFFFF" class="title"> 
      <td width="32" align="center"><strong>选择</strong></td>
      <td width="31" align="center"><strong>ID</strong></td>
      <td width="95" height="22" align="center"><strong> 模板名称</strong></td>
      <td width="129" align="center"><strong>效果图</strong></td>
      <td width="64" align="center"><strong>设计者</strong></td>
      <td width="82" align="center"><strong>模板类型</strong></td>
      <td width="131" height="22" align="center"><strong> 操作</strong></td>
    </tr>
    <%if not(rs.bof and rs.eof) then
  do while not rs.EOF %>
    <tr bgcolor="#FFFFFF" class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''"> 
      <td width="32" align="center"> 
        <input type="radio" value="<%=rs("SkinID")%>" <%if rs("IsDefault")=true then response.write " checked"%> name="SkinID"></td>
      <td width="31" align="center"><%=rs("SkinID")%></td>
      <td align="center"><%=rs("SkinName")%></td>
      <td width="129" align="center"> 
        <%response.write "<a href='Admin_Skin.asp?Action=Prview&SkinID=" & rs("SkinID") & "' title='点此查看原始效果图'><img src='" & rs("PicUrl") & "' width=100 height=30 border=0></a>"%></td>
      <td width="64" align="center"> 
        <%response.write "<a href='mailto:" & rs("DesignerEmail") & "' title='设计者信箱：" & rs("DesignerEmail") & vbcrlf & "设计者主页：" & rs("DesignerHomepage") & "'>" & rs("DesignerName") & "</a>"%></td>
      <td width="82" align="center"> 
        <%if rs("DesignType")=1 then response.write "用户自定义" else response.write "系统提供"%></td>
      <td width="131" align="center"> 
        <%
	response.write "<a href='Admin_Skin.asp?Action=Modify&SkinID=" & rs("SkinID") & "'>修改模板设置</a>&nbsp;"
	if rs("DesignType")=1 and rs("IsDefault")=False then
		response.write "<a href='Admin_Skin.asp?Action=Del&SkinID=" & rs("SkinID") & "' onClick=""return confirm('确定要删除此配色模板吗？删除此配色模板后原使用此配色模板的文章将改为使用系统默认配色模板。');"">删除模板</a>"
	else
		response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
	end if
	%> </td>
    </tr>
    <%
		rs.MoveNext
   	loop
  %>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan="7" align="center"> 
        <input name="Action" type="hidden" id="Action" value="Set"> 
        <input type="submit" name="Submit" value="将选中的模板设为默认模板"></td>
    </tr>
    <%end if%>
  </table>  
</form>
<%
	rs.close
	set rs=nothing
end sub

sub Export()
	sql="select * from Skin"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,1
%>
<form name="myform" method="post" action="Admin_Skin.asp">
  <table width="600" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="border">
    <tr bgcolor="#FFFFFF" class="topbg"> 
      <td height="22" colspan="6" align="center"><strong>配色模板导出</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="title"> 
      <td width="31" align="center"><strong>选择</strong></td>
      <td width="52" align="center"><strong>ID</strong></td>
      <td width="140" height="22" align="center"><strong> 模板名称</strong></td>
      <td width="158" align="center"><strong>效果图</strong></td>
      <td width="104" align="center"><strong>设计者</strong></td>
      <td width="84" height="22" align="center"><strong>模板类型</strong></td>
    </tr>
    <%if not(rs.bof and rs.eof) then
  do while not rs.EOF %>
    <tr bgcolor="#FFFFFF" class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''"> 
      <td width="31" align="center"> 
        <input type="checkbox" value="<%=rs("SkinID")%>" name="SkinID" onclick="unselectall()"></td>
      <td width="52" align="center"><%=rs("SkinID")%></td>
      <td align="center"><%=rs("SkinName")%></td>
      <td width="158" align="center"> 
        <%response.write "<a href='Admin_Skin.asp?Action=Prview&SkinID=" & rs("SkinID") & "' title='点此查看原始效果图'><img src='" & rs("PicUrl") & "' width=100 height=30 border=0></a>"%></td>
      <td width="104" align="center"> 
        <%response.write "<a href='mailto:" & rs("DesignerEmail") & "' title='设计者信箱：" & rs("DesignerEmail") & vbcrlf & "设计者主页：" & rs("DesignerHomepage") & "'>" & rs("DesignerName") & "</a>"%></td>
      <td width="84" align="center"> 
        <%if rs("DesignType")=1 then response.write "用户自定义" else response.write "系统提供"%></td>
    </tr>
    <%
		rs.MoveNext
   	loop
  %>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan="6"> 
        <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
        选中所有模板&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;导出选中的模板到数据库： 
        <input name="SkinMdb" type="text" id="SkinMdb" value="Skin/Skin.mdb" size="20" maxlength="50">
        <input type="submit" name="Submit" value="导出">
        <input name="Action" type="hidden" id="Action" value="DoExport"></td>
    </tr>
    <%end if%>
  </table>  
</form>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled!=true)
       e.checked = form.chkAll.checked;
    }
}
</script>
<%
	rs.close
	set rs=nothing
end sub

sub Import()
%>
<form name="myform" method="post" action="Admin_Skin.asp">
  <table width="600" border="0" align="center" cellpadding="2" cellspacing="1" class="border">
    <tr class="topbg"> 
      <td height="22" align="center"><strong>配色模板导入（第一步）</strong></td>
    </tr>
	<tr class="tdbg">
      <td height="100">&nbsp;&nbsp;&nbsp;&nbsp;请输入要导入的模板数据库的文件名： 
        <input name="SkinMdb" type="text" id="SkinMdb" value="Skin/Skin.mdb" size="20" maxlength="50">
        <input name="Submit" type="submit" id="Submit" value=" 下一步 ">
        <input name="Action" type="hidden" id="Action" value="Import2"> </td>
	</tr>
  </table>
</form>
<%
end sub

sub Import2()
	on error resume next
	dim mdbname,tconn,trs
	mdbname=replace(trim(request.form("skinmdb")),"'","")
	if mdbname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请填写导入模版数据库名"
		exit sub
	end if
	
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	if err.number<>0 then
		ErrMsg=ErrMsg & "<br><li>数据库操作失败，请以后再试，错误原因：" & err.Description
		err.clear
		exit sub
	end if
	

	sql="select * from Skin"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,tconn,1,1
%>
<form name="myform" method="post" action="Admin_Skin.asp">
  <table width="600" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="border">
    <tr bgcolor="#FFFFFF" class="topbg"> 
      <td height="22" colspan="6" align="center"><strong>配色模板导入（第二步）</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="title"> 
      <td width="39" align="center"><strong>选择</strong></td>
      <td width="41" align="center"><strong>ID</strong></td>
      <td width="159" height="22" align="center"><strong> 模板名称</strong></td>
      <td width="150" align="center"><strong>效果图</strong></td>
      <td width="100" align="center"><strong>设计者</strong></td>
      <td width="80" height="22" align="center"><strong>模板类型</strong></td>
    </tr>
    <%if not(rs.bof and rs.eof) then
  do while not rs.EOF %>
    <tr bgcolor="#FFFFFF" class="tdbg" onmouseover="this.style.backgroundColor='#BFDFFF'" onmouseout="this.style.backgroundColor=''"> 
      <td width="39" align="center"> 
        <input type="checkbox" value="<%=rs("SkinID")%>" name="SkinID" onclick="unselectall()"></td>
      <td width="41" align="center"><%=rs("SkinID")%></td>
      <td align="center"><%=rs("SkinName")%></td>
      <td width="150" align="center"> 
        <%response.write "<a href='Admin_Skin.asp?Action=Prview&SkinID=" & rs("SkinID") & "' title='点此查看原始效果图'><img src='" & rs("PicUrl") & "' width=100 height=30 border=0></a>"%></td>
      <td width="100" align="center"> 
        <%response.write "<a href='mailto:" & rs("DesignerEmail") & "' title='设计者信箱：" & rs("DesignerEmail") & vbcrlf & "设计者主页：" & rs("DesignerHomepage") & "'>" & rs("DesignerName") & "</a>"%></td>
      <td width="80" align="center"> 
        <%if rs("DesignType")=1 then response.write "用户自定义" else response.write "系统提供"%></td>
    </tr>
    <%
		rs.MoveNext
   	loop
  %>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan="6"> 
        <input name="chkAll" type="checkbox" id="chkAll" onclick=CheckAll(this.form) value="checkbox">
        选中所有模板&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="submit" name="Submit" value="导入选中的模板">
        <input name="SkinMdb" type="hidden" id="SkinMdb" value="<%=mdbname%>">
        <input name="Action" type="hidden" id="Action" value="DoImport"></td>
    </tr>
    <%end if%>
  </table>  
</form>
<SCRIPT language=javascript>
function unselectall()
{
    if(document.myform.chkAll.checked){
	document.myform.chkAll.checked = document.myform.chkAll.checked&0;
    } 	
}

function CheckAll(form)
{
  for (var i=0;i<form.elements.length;i++)
    {
    var e = form.elements[i];
    if (e.Name != "chkAll"&&e.disabled!=true)
       e.checked = form.chkAll.checked;
    }
}
</script>
<%
	rs.close
	set rs=nothing
end sub

sub ShowSkinSet()
	if Action="Add" then
		sql="select * from Skin where IsDefault=True"
	elseif Action="Modify" then
		if SkinID="" then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>请指定SkinID</li>"
			exit sub
		else
			SkinID=Clng(SkinID)
		end if
		sql="select * from Skin where SkinID=" & SkinID
	end if
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,1
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>数据库出现错误！</li>"
		rs.close
		set rs=nothing
		exit sub
	end if
	dim Skin_CSS
	Skin_CSS=split(rs("Skin_CSS"),"|||")
%>
<form name="form1" method="post" action="Admin_Skin.asp">
  <table width="600" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" class="border">
    <tr align="center" bgcolor="#FFFFFF" class="title"> 
      <td height="22" colspan="2"><strong> 
        <%if Action="Add" then%>
        添加新配色模板 
        <%else%>
        修改模板设置 
        <%end if%>
        </strong></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="topbg"> 
      <td height="20" colspan="2"><strong>配色模板基本信息</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="40%"><strong>配色模板名称：</strong></td>
      <td> 
        <input name="SkinName" type="text" id="SkinName" value="<%if Action="Modify" then response.write rs("SkinName")%>" size="50" maxlength="50"> 
        <input name="SkinID" type="hidden" id="SkinID" value="<%=SkinID%>"></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="40%"><strong>配色模板预览图：</strong></td>
      <td> 
        <%if Action="Modify" and rs("DesignType")=0 then%> <input name="PicUrl" type="hidden" id="PicUrl" value="<%=rs("PicUrl")%>"> 
        <%=rs("PicUrl")%> <%else%> <input name="PicUrl" type="text" id="PicUrl" value="<%=rs("PicUrl")%>" size="50" maxlength="100"> 
        <%end if%></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="40%"><strong>设计者姓名：</strong></td>
      <td> 
        <%if Action="Modify" and rs("DesignType")=0 then%> <input name="DesignerName" type="hidden" id="DesignerName" value="<%=rs("DesignerName")%>"> 
        <%=rs("DesignerName")%> <%else%> <input name="DesignerName" type="text" id="DesignerName" value="<%=rs("DesignerName")%>" size="50" maxlength="30"> 
        <%end if%></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="40%"><strong>设计者Email：</strong></td>
      <td> 
        <%if Action="Modify" and rs("DesignType")=0 then%> <input name="DesignerEmail" type="hidden" id="DesignerEmail" value="<%=rs("DesignerEmail")%>"> 
        <%=rs("DesignerEmail")%> <%else%> <input name="DesignerEmail" type="text" id="DesignerEmail" value="<%=rs("DesignerEmail")%>" size="50" maxlength="100"> 
        <%end if%></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td width="40%"><strong>设计者主页：</strong></td>
      <td> 
        <%if Action="Modify" and rs("DesignType")=0 then%> <input name="DesignerHomepage" type="hidden" id="DesignerHomepage" value="<%=rs("DesignerHomepage")%>"> 
        <%=rs("DesignerHomepage")%> <%else%> <input name="DesignerHomepage" type="text" id="DesignerHomepage" value="<%=rs("DesignerHomepage")%>" size="50" maxlength="100"> 
        <%end if%></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="topbg"> 
      <td height="20" colspan="2"><strong>模板配色设置（修改以下设置必须具备一定网页知识，<font color="#FFFF00">不能使用单引号或双引号，否则会容易造成程序错误</font>）</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>BODY标签</strong><br>
        控制整个页面风格的背景颜色或者背景图片等</td>
      <td> 
        <input name="Body" type="text" id="Body" value="<%=rs("Body")%>" size="50" maxlength="200"></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">链接总的CSS定义</font></strong><br>
        可定义内容为链接字体颜色、样式等</td>
      <td> 
        <textarea name="Link" cols="41" rows="4" id="Link"><%=Skin_CSS(0)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">Body的CSS定义</font></strong><br>
        对应CSS中“BODY”，可定义内容为网页字体颜色、背景、浏览器边框等</td>
      <td> 
        <textarea name="CSS_Body" cols="41" rows="4" id="CSS_Body"><%=Skin_CSS(1)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">单元格的CSS定义</font></strong><br>
        对应CSS中的“TD”，这里为总的表格定义，为一般表格的的单元格风格设置，可定义内容为背景、字体颜色、样式等</td>
      <td> 
        <textarea name="TD" cols="41" rows="4" id="TD"><%=Skin_CSS(2)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td> 
        <p><strong><font color="#FF0000">文本框的CSS定义</font></strong><br>
          对应CSS中的“INPUT”，这里为文本框的风格设置，可定义内容为背景、字体、颜色、边框等</p></td>
      <td> 
        <textarea name="Input" cols="41" rows="4" id="Input"><%=Skin_CSS(3)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">按钮的CSS定义</font></strong><br>
        对应CSS中的“BUTTON”，这里为按钮的风格设置，可定义内容为背景、字体、颜色、边框等</td>
      <td> 
        <textarea name="Button" cols="41" rows="4" id="Button"><%=Skin_CSS(4)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">下拉列表框的CSS定义</font></strong><br>
        对应CSS中的“SELECT”，这里为下拉列表框的风格设置，可定义内容为背景、字体、颜色、边框等 </td>
      <td> 
        <textarea name="Select" cols="41" rows="5" id="Select"><%=Skin_CSS(5)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">表格边框的CSS定义</font></strong><font color="#FF0000">一</font><br>
        对应CSS中的“.border”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=border 
        </font></td>
      <td> 
        <textarea name="Border" cols="41" rows="5" id="Border"><%=Skin_CSS(6)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>表格边框的CSS定义</strong>二<br>
        对应CSS中的“.border2”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=border2</font></td>
      <td> 
        <textarea name="Border2" cols="41" rows="5" id="Border2"><%=Skin_CSS(7)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td bgcolor="#FFFFFF"><strong><font color="#FF0000">标题文字的CSS定义</font></strong><br>
        对应CSS中的“.FootBg”，可定义内容为背景、背景图、字体及其颜色等<br>
        <font color="#0000FF">调用：Class=FootBg</font></td>
      <td> 
        <textarea name="FootBg" cols="41" rows="4" id="FootBg"><%=Skin_CSS(8)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">导航条菜单的CSS定义</font></strong><br>
        对应CSS中的“.title”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=title 
        </font></td>
      <td> 
        <textarea name="Title" cols="41" rows="4" id="textarea3"><%=Skin_CSS(9)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong><font color="#FF0000">内容单元格的CSS定义</font></strong><br>
        对应CSS中的“.tdbg”，可定义内容为背景、背景图、字体及其颜色等(<font color="#FF0000">首页产品列表</font>)<br> <font color="#0000FF">调用：Class=tdbg</font></td>
      <td> 
        <textarea name="tdbg" cols="41" rows="4" id="textarea"><%=Skin_CSS(10)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>导航条内容的CSS定义</strong><br>
        对应CSS中的“.txt_css”，可定义内容为字体、颜色等<br> <font color="#0000FF">调用：Class=txt_css</font></td>
      <td> 
        <textarea name="txt_css" cols="41" rows="4" id="textarea2"><%=Skin_CSS(11)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan="2"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr> 
            <td><strong><font color="#CC0000">［左分栏格式表格的CSS定义］</font></strong></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;左分栏－标题文字的CSS定义</strong><br>
        对应CSS中的“.title_lefttxt”，可定义内容为字体、颜色等<br> <font color="#0000FF">调用：Class=title_lefttxt</font></td>
      <td> 
        <textarea name="title_lefttxt" cols="41" rows="4" id="textarea4"><%=Skin_CSS(12)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;<font color="#FF0000">左分栏－标题单元格的CSS定义（1）</font></strong><br>
        对应CSS中的“.title_left”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=title_left</font></td>
      <td> 
        <textarea name="Title_Left" cols="41" rows="4" id="Title_Left"><%=Skin_CSS(13)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;<font color="#FF0000">左分栏－内容单元格的CSS定义（1）</font></strong><br>
        对应CSS中的“.tdbg_left”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=tdbg_left</font></td>
      <td> 
        <textarea name="tdbg_left" cols="41" rows="4" id="textarea5"><%=Skin_CSS(14)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;左分栏－标题单元格的CSS定义（2）</strong><br>
        对应CSS中的“.title_left2”，可定义内容为背景、背景图、字体及其颜色等<font color="#666666">（注：现只为绿雨飘香模板中使用）</font><br> 
        <font color="#0000FF">调用：Class=title_left2</font></td>
      <td> 
        <textarea name="Title_Left2" cols="41" rows="4" id="Title_Left2"><%=Skin_CSS(15)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;左分栏－内容单元格的CSS定义（2）</strong><br>
        对应CSS中的“.tdbg_left2”，可定义内容为背景、背景图、字体及其颜色等<font color="#666666">（注：现为灰色阴影条的CSS）</font><br> 
        <font color="#0000FF">调用：Class=tdbg_left2</font></td>
      <td> 
        <textarea name="tdbg_left2" cols="41" rows="4" id="textarea6"><%=Skin_CSS(16)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;左分栏－内容单元格背景的CSS定义</strong><br>
        对应CSS中的“.tdbg_leftall”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=tdbg_leftall</font></td>
      <td> 
        <textarea name="tdbg_leftall" cols="41" rows="4" id="textarea7"><%=Skin_CSS(17)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td colspan="2"> 
        <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
          <tr> 
            <td><strong><font color="#CC0000">［右分栏格式表格的CSS定义］</font></strong></td>
          </tr>
        </table></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;右分栏－标题文字的CSS定义</strong><br>
        对应CSS中的“.title_righttxt”，可定义内容为字体、颜色等<br> <font color="#0000FF">调用：Class=title_righttxt</font></td>
      <td> 
        <textarea name="title_righttxt" cols="41" rows="4" id="textarea8"><%=Skin_CSS(18)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;<font color="#FF0000">右分栏－标题单元格的CSS定义（1）</font></strong><br>
        对应CSS中的“.title_right”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=title_right</font></td>
      <td> 
        <textarea name="Title_Right" cols="41" rows="4" id="Title_Right"><%=Skin_CSS(19)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;右分栏－内容单元格的CSS定义（1）</strong><br>
        对应CSS中的“.tdbg_right”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=tdbg_right</font></td>
      <td> 
        <textarea name="tdbg_right" cols="41" rows="4" id="textarea17"><%=Skin_CSS(20)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;右分栏－标题单元格的CSS定义（2）</strong><br>
        对应CSS中的“.title_right2”，可定义内容为背景、背景图、字体及其颜色等<font color="#666666">（注：备用的CSS）</font><br> 
        <font color="#0000FF">调用：Class=title_right2</font></td>
      <td> 
        <textarea name="Title_Right2" cols="41" rows="4" id="textarea16"><%=Skin_CSS(21)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;右分栏－内容单元格的CSS定义（2）</strong><br>
        对应CSS中的“.tdbg_right2”，可定义内容为背景、背景图、字体及其颜色等<font color="#666666">（注：备用的CSS）</font><br> 
        <font color="#0000FF">调用：Class=tdbg_right2</font></td>
      <td> 
        <textarea name="tdbg_right2" cols="41" rows="4" id="textarea18"><%=Skin_CSS(22)%></textarea></td>
    </tr>
    <tr bgcolor="#FFFFFF" class="tdbg"> 
      <td><strong>&middot;右分栏－内容单元格背景的CSS定义</strong><br>
        对应CSS中的“.tdbg_rightall”，可定义内容为背景、背景图、字体及其颜色等<br> <font color="#0000FF">调用：Class=tdbg_rightall</font></td>
      <td> 
        <textarea name="tdbg_rightall" cols="41" rows="4" id="textarea10"><%=Skin_CSS(23)%></textarea></td>
    </tr>
    <tr align="center" bgcolor="#FFFFFF" class="tdbg"> 
      <td height="50" colspan="2"> 
        <%if Action="Add" then%> <input name="Action" type="hidden" id="Action" value="SaveAdd"> 
        <input type="submit" name="Submit2" value=" 添 加 "> <%else%> <input name="Action" type="hidden" id="Action" value="SaveModify"> 
        <input type="submit" name="Submit2" value=" 保存修改结果 "> <%end if%> </td>
    </tr>
  </table>
</form>
<%
	rs.close
	set rs=nothing
end sub
%>
<!-- #include file="Inc/Foot.asp" -->
<%
sub SaveAdd()
	call CheckSkin()
	if FoundErr=True then exit sub
	
	sql="select top 1 * from Skin"
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	rs.addnew
	rs("IsDefault")=False
	rs("DesignType")=1
	call SaveSkin()
	rs.close
	set rs=nothing
	call WriteSuccessMsg("成功添加新的配色模板："& trim(request("SkinName")),"")	
end sub

sub SaveModify()
	
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定SkinID</li>"
	else
		SkinID=Clng(SkinID)
	end if
	call CheckSkin()
	if FoundErr=True then exit sub
	
	sql="select * from Skin where SkinID=" & SkinID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的配色模板！</li>"
	else
		call SaveSkin()
		call WriteSuccessMsg("保存配色模板设置成功！","")
	end if
	rs.close
	set rs=nothing	
end sub

sub SetDefault()
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定SkinID</li>"
		exit sub
	else
		SkinID=Clng(SkinID)
	end if
	conn.execute("update Skin set IsDefault=False where IsDefault=True")
	conn.execute("update Skin set IsDefault=True where SkinID=" & SkinID)
	call WriteSuccessMsg("成功将选定的模板设置为默认模板","")
end sub

sub CheckSkin()
	if trim(request("SkinName"))="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>模板名称不能为空！</li>"
	end if
	if trim(request("PicUrl"))="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>模板预览图地址不能为空！</li>"
	end if
	if trim(request("DesignerName"))="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>模板设计者姓名不能为空！</li>"
	end if
	if trim(request("DesignerEmail"))="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>模板设计者邮箱不能为空！</li>"
	end if
end sub

sub SaveSkin()
	rs("SkinName")=trim(request("SkinName"))
	rs("PicUrl")=trim(request("PicUrl"))
	rs("DesignerName")=trim(request("DesignerName"))
	rs("DesignerEmail")=trim(request("DesignerEmail"))
	rs("DesignerHomePage")=trim(request("DesignerHomepage"))
	rs("Body")=trim(request("Body"))
	dim Skin_CSS
		Skin_CSS= request("Link") & "|||" & request("CSS_Body") & "|||" & request("TD") & "|||" & request("Input") & "|||" & request("Button") & "|||" & request("Select") & "|||"
	Skin_CSS=Skin_CSS & request("border") & "|||" & request("border2") & "|||" & request("FootBg") & "|||" & request("title") & "|||" & request("tdbg") & "|||" & request("txt_css") & "|||"
	Skin_CSS=Skin_CSS & request("title_lefttxt") & "|||" & request("title_left") & "|||" & request("tdbg_left") & "|||" & request("title_left2") & "|||" & request("tdbg_left2") & "|||" & request("tdbg_leftall") & "|||"	
	Skin_CSS=Skin_CSS & request("title_righttxt") & "|||" & request("title_right") & "|||" & request("tdbg_right") & "|||" & request("title_right2") & "|||" & request("tdbg_right2") & "|||" & request("tdbg_rightall")	
	rs("Skin_CSS")=Skin_CSS
	rs.update
end sub

sub DelSkin()
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定SkinID</li>"
		exit sub
	else
		SkinID=Clng(SkinID)
	end if
	sql="select * from Skin where SkinID=" & SkinID
	Set rs=Server.CreateObject("Adodb.RecordSet")
	rs.Open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>找不到指定的配色模板！</li>"
	else
		if rs("DesignType")=0 then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>不能删除系统自带的模板！</li>"
		elseif rs("IsDefault")=True then
			FoundErr=True
			ErrMsg=ErrMsg & "<br><li>当前模板为默认模板，不能删除。请先将默认模板改为其他模板后再来删除此模板。</li>"
		end if
	end if
	if FoundErr=True then
		rs.close
		set rs=nothing
		exit sub
	end if
	rs.delete
	rs.update
	rs.close
	set rs=nothing
	dim trs
	set trs=conn.execute("select SkinID from Skin where IsDefault=True")
	conn.execute("update ArticleClass set SkinID=" & trs(0) & " where SkinID=" & SkinID)
	conn.execute("update Article set SkinID=" & trs(0) & " where SkinID=" & SkinID)
	set trs=nothing
	call WriteSuccessMsg("成功删除选定的模板。并将使用此模板的栏目和文章改为使用默认模板。","")	
end sub

sub DoExport()
	on error resume next
	dim mdbname,tconn,trs
	SkinID=replace(SkinID,"'","")
	mdbname=replace(trim(request.form("skinmdb")),"'","")
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要导出的模版</li>"
	end if
	if mdbname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请填写导出模版数据库名"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	if err.number<>0 then
		ErrMsg=ErrMsg & "<br><li>数据库操作失败，请以后再试，错误原因：" & err.Description
		err.clear
		exit sub
	end if
	tconn.execute("delete * from Skin")
	set rs=conn.execute("select * from Skin where SkinID in (" & SkinID &")  order by SkinID ")
	set trs=server.CreateObject("adodb.recordset")
	trs.open "select * from Skin",tconn,1,3
	do while not rs.eof
		trs.addnew
		trs("SkinName")=rs("SkinName")
		trs("PicUrl")=rs("PicUrl")
		trs("DesignerName")=rs("DesignerName")
		trs("DesignerEmail")=rs("DesignerEmail")
		trs("DesignerHomePage")=rs("DesignerHomePage")
		trs("Body")=rs("Body")
		trs("Skin_CSS")=rs("Skin_CSS")
		trs("IsDefault")=False
		trs("DesignType")=rs("DesignType")
		trs.update
		rs.movenext
	loop
	trs.close
	set trs=nothing
	rs.close
	set rs=nothing
	tconn.close
	set tconn=nothing
	call WriteSuccessMsg("已经成功将所选中的模板设置导出到指定的数据库中！<br><br>你还需要将Skin文件夹中图片文件一起打包。","")
end sub

sub DoImport()
	on error resume next
	dim mdbname,tconn,trs
	SkinID=replace(SkinID,"'","")
	mdbname=replace(trim(request.form("skinmdb")),"'","")
	if SkinID="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请指定要导入的模版</li>"
	end if
	if mdbname="" then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>请填写导出模版数据库名"
	end if
	
	if FoundErr=True then
		exit sub
	end if
	
	Set tconn = Server.CreateObject("ADODB.Connection")
	tconn.open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(mdbname)
	if err.number<>0 then
		ErrMsg=ErrMsg & "<br><li>数据库操作失败，请以后再试，错误原因：" & err.Description
		err.clear
		exit sub
	end if
	
	set rs=tconn.execute(" select * from Skin where SkinID in (" & SkinID &")  order by SkinID")
	set trs=server.CreateObject("adodb.recordset")
	trs.open "select * from Skin",conn,1,3
	do while not rs.eof
		trs.addnew
		trs("SkinName")=rs("SkinName")
		trs("PicUrl")=rs("PicUrl")
		trs("DesignerName")=rs("DesignerName")
		trs("DesignerEmail")=rs("DesignerEmail")
		trs("DesignerHomePage")=rs("DesignerHomePage")
		trs("Body")=rs("Body")
		trs("Skin_CSS")=rs("Skin_CSS")
		trs("IsDefault")=False
		trs("DesignType")=rs("DesignType")
		trs.update
		rs.movenext
	loop
	trs.close
	set trs=nothing
	rs.close
	set rs=nothing
	tconn.close
	set tconn=nothing
	call WriteSuccessMsg("已经成功从指定的数据库中导入选中的模板！<br><br>你还需要将图片文件复制到Skin目录中的相应文件夹中才真正完成导入工作。","")
end sub
%>
