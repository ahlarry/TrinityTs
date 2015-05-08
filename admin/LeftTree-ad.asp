<%
'=================================
'
'    QQ:290116505
'    http://www.wrtx.cn/qywz2
'    copyright(c)2006 网软天下企业网站管理系统超漂亮版
'
'=================================
%>
<html><head>
<meta http-equiv=Content-Type content=text/html;charset=gb2312>
<link href="Inc/southidc.css" rel="stylesheet" type="text/css">
<style type="text/css">
BODY {
	MARGIN: 0px;
	background-position: 799;
	background-color: #975102;
}

.sec_menu {
	BORDER-RIGHT: white 1px solid;
	OVERFLOW: hidden;
	BORDER-LEFT: white 1px solid;
	BORDER-BOTTOM: white 1px solid;
	background-color: #FFF1CE;
}

.menu_title SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #993300; POSITION: relative; TOP: 2px 
}

.menu_title2 SPAN {
	FONT-WEIGHT: bold; LEFT: 10px; COLOR: #CC6600; POSITION: relative; TOP: 2px
}
.STYLE1 {
	color: #990000;
	font-weight: bold;
}
.STYLE2 {color: #990000}
</style>
</head>

<BODY>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		<td valign="bottom" height="42">
		<img height="38" src="image/title.gif" width="158" border="0"></td>
	</tr>
	<tr>
		
    <td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="image/title_bg_quit.gif" height="25"> 
      <span><span class="STYLE1"><a href="../index.asp" target="_top" class="STYLE2">回到首页</a>
      | </span> <a target="_top" href="Loginout.asp"><font color="215DC6"><b></font><span class="STYLE2">退出</span> </a></span></td>
	</tr>
	<tr>
		
    <td align="center" onMouseOver="aa('up')" onMouseOut="StopScroll()">&nbsp; </td>
	</tr>
</table>
<script>
var he=document.body.clientHeight-105
document.write("<div id=tt style=height:"+he+";overflow:hidden>")
</script>
<table cellspacing="0" cellpadding="0" width="160" align="center">
	<tr>
		
    <td width="158" height="25" background=image/menudown.gif class="menu_title" id="imgmenu1" style="cursor:hand" onClick="showsubmenu(1)" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';"> 
      <span>系统管理 </span></td>
	</tr>

	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a target="main" href="Admin_Manage.asp">管理员管理</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Admin_AddAffiche.asp">网站公告</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="SiteConfig.asp">网站配置</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Bs.asp"></a><a target="main" href="Admin_DataBackup.asp">数据库备份</a></td>
          </tr>
          <tr> 
            <td><a href="../Count/index.asp" target="_blank">网站统计系统</a></td>
          </tr>
          <tr> 
            <td><a href="SouthidcEditor/Admin_Style.asp" target="main">在线编辑器管理</a></td>
          </tr>
          <tr> 
            <td><a href="aspcheck.asp" target="main">ASP服务器探针</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Admin_UploadFileManage.asp">上传文件管理</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Help.asp">系统帮<font color="000000">助</font></a></td>
          </tr>
        </table>
      </div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu2" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(2)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>企业信息 </span></td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a href="AdminAboutusAdd.asp" target="main">新增企业信息</a></td>
          </tr>
          <tr> 
            <td><a href="AdminAboutus.asp" target="main">管理企业信息</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_Culture.asp" target="main">管理企业文化</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_CultureNewsAdd.asp" target="main">增加企业文化</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu3" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(3)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>产品管理 </span></td>
	</tr>
	<tr>
		<td id="submenu3" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="ClassManage.asp"><font color="000000">产品类别</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="ProductManage.asp"><font color="000000">产品管理</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="ProductAdd.asp"><font color="000000">添加产品</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="ProductCheck.asp"><font color="000000">审核产品</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu4" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(4)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>询价管理 </span></td>
	</tr>
	<tr>
		<td id="submenu4" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a target="main" href="Admin_Order.asp"><font color="000000">询价管理</font></a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>		
    <td id="imgmenu14" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(14)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>下载中心 </span></td>
	</tr>
	<tr>
		<td id="submenu14" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a href="Down_add.asp" target="main">添加下载程序</a></td>
          </tr>
          <tr> 
            <td><a href="Down_Manage.asp" target="main">管理下载程序</a></td>
          </tr>
          <tr> 
            <td><a href="Down_ClassManage.asp" target="main">下载程序类别</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>


	<tr>
		
    <td id="imgmenu5" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(5)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>会员管理 </span></td>
	</tr>
	<tr>
		<td id="submenu5" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="UserManage.asp"><font color="000000">注册会员管理</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu6" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(6)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>新闻管理/FAQ </span></td>
	</tr>
	<tr>
		<td id="submenu6" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td> <a href="News_add.asp" target="main">添加新闻内容</a> </td>
          </tr>
          <tr> 
            <td><a href="News_Manage.asp" target="main">管理全部新闻</a></td>
          </tr>
          <tr> 
            <td><a href="News_ClassManage.asp" target="main">管理新闻类别</a></td>
          </tr>
          <tr>
            <td><a href="Admin_faq.asp" target="main">常见问题</a></td>
          </tr>
          <tr>
            <td><a href="Admin_faqNewsAdd.asp" target="main">添加常见问题</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>
	
	<tr>
		
    <td id="imgmenu7" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(7)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>留言管理 </span></td>
	</tr>
	<tr>
		<td id="submenu7" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_Feedback.asp"><font color="000000">留言管理</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_FeedbackAdd.asp"><font color="000000">管理员公告</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu8" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(8)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>荣誉管理 </span></td>
	</tr>
	<tr>
		<td id="submenu8" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_Honor.asp"><font color="000000">企业荣誉管理</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HonorAdd.asp"><font color="000000">添加企业荣誉</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_CompVisualize.asp"><font color="000000">企业形象管理</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_CompVisualizeAdd.asp"><font color="000000">添加企业形象</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu10" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(10)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>营销网络/客户 </span></td>
	</tr>
	<tr>
		<td id="submenu10" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
		  <table cellspacing="3" cellpadding="0" width="130" align="center">
            <tr>
              <td><a target="main" href="market_input.asp"><font color="000000">增加国内营销网络</font></a></td>
            </tr>
            <tr>
              <td><a target="main" href="market.asp">管理国内营销网络</a></td>
            </tr>
            <tr>
              <td><a target="main" href="Admin_OverseasMarket.asp"><font color="000000">国际市场</font></a></td>
            </tr>
            <tr>
              <td><a href="class_customer.asp" target="main">管理客户类别</a></td>
            </tr>
            <tr>
              <td><a target="main" href="zp_input.asp">增加客户</a></td>
            </tr>
            <tr>
              <td><a target="main" href="zp.asp">管理客户</a></td>
            </tr>
          </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu9" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(9)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>人才管理 </span></td>
	</tr>
	<tr>
		<td id="submenu9" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_HrDemand.asp"><font color="000000">招聘管理</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HrDemandAdd.asp"><font color="000000">发布招聘</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HrManage.asp"><font color="000000">应聘管理</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HrPolicy.asp"><font color="000000">人才策略</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu11" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(11)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>调查管理 </span></td>
	</tr>
	<tr>
		<td id="submenu11" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="VoteManage.asp"><font color="000000">调查设置</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="VoteAdd.asp"><font color="000000">添加新调查</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>		
    <td id="imgmenu12" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(12)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>邮件列表 </span></td>
	</tr>
	<tr>
		<td id="submenu12" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a target="main" href="AdminMaillist.asp"><font color="000000">邮件列表</font></a></td>
          </tr>
          <tr> 
            <td><a href="AdminMaillist.asp?Action=Export" target="main">列表导出</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>		
    <td id="imgmenu13" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(13)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>友情链接 </span></td>
	</tr>
	<tr>
		<td id="submenu13" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_FriendLinks.asp"><font color="000000">友情链接管理</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

<tr>		
    <td id="imgmenu15" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(15)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>模板管理 </span></td>
  </tr>
	<tr>
		<td id="submenu15" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a href="Admin_Skin.asp" target="main">配色模板管理</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_Skin.asp?Action=Add" target="main">配色模板添加</a></td>
          </tr>
          <tr>
            <td><a href="Admin_Skin.asp?Action=Export" target="main">配色模板导出</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_Skin.asp?Action=Import" target="main">配色模板导入</a></td>
          </tr>
        </table>
      </div>
		<br>
		</td>
	</tr>


</table>
&nbsp;
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td class="menu_title" onMouseOver="this.className='menu_title2';" onMouseOut="this.className='menu_title';" background="image/title_bg_quit.gif" height="25"> 
      <span>版权所有</span> </td>
	</tr>
	<tr>
		<td>
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
			<table cellspacing="4" cellpadding="3">
				<tr>
					
              <td width="141" height="20" style="line-height: 150%;">
                Copyright：<br><a target="_blank" href="/"><font color="000000"
					>nanidc.com</font></a><br>
				  </td>
				</tr>
			</table>
			</div>
		</div>
		</td>
	</tr>
</table>
</div>
<table cellspacing="0" cellpadding="0" width="158" align="center">
	<tr>
		
    <td align="center" onMouseOver="aa('Down')" onMouseOut="StopScroll()" valign="bottom">&nbsp; 
    </td>
	</tr>
</table>
<script>

function aa(Dir)
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//这里100为滚动速度
function StopScroll(){if(Timer!=null)clearTimeout(Timer)}

function initIt(){
divColl=document.all.tags("DIV");
for(i=0; i<divColl.length; i++) {
whichEl=divColl(i);
if(whichEl.className=="child")whichEl.style.display="none";}
}
function expands(el) {
whichEl1=eval(el+"Child");
if (whichEl1.style.display=="none"){
initIt();
whichEl1.style.display="block";
}else{whichEl1.style.display="none";}
}
var tree= 0;
function loadThreadFollow(){
if (tree==0){
document.frames["hiddenframe"].location.replace("LeftTree.asp");
tree=1
}
}

function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
imgmenu = eval("imgmenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
imgmenu.background="image/menuup.gif";
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
imgmenu.background="image/menudown.gif";
}
}

function loadingmenu(id){
var loadmenu =eval("menu" + id);
if (loadmenu.innerText=="Loading..."){
document.frames["hiddenframe"].location.replace("LeftTree.asp?menu=menu&id="+id+"");
}
}
top.document.title="网软天下企业网站管理系统超漂亮版企业网站管理系统"; 
</script>
</BODY></HTML>
