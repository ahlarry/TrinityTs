<%
'=================================
'
'    QQ:290116505
'    http://www.wrtx.cn/qywz2
'    copyright(c)2006 ����������ҵ��վ����ϵͳ��Ư����
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
      <span><span class="STYLE1"><a href="../index.asp" target="_top" class="STYLE2">�ص���ҳ</a>
      | </span> <a target="_top" href="Loginout.asp"><font color="215DC6"><b></font><span class="STYLE2">�˳�</span> </a></span></td>
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
      <span>ϵͳ���� </span></td>
	</tr>

	<tr>
		<td id="submenu1" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a target="main" href="Admin_Manage.asp">����Ա����</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Admin_AddAffiche.asp">��վ����</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="SiteConfig.asp">��վ����</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Bs.asp"></a><a target="main" href="Admin_DataBackup.asp">���ݿⱸ��</a></td>
          </tr>
          <tr> 
            <td><a href="../Count/index.asp" target="_blank">��վͳ��ϵͳ</a></td>
          </tr>
          <tr> 
            <td><a href="SouthidcEditor/Admin_Style.asp" target="main">���߱༭������</a></td>
          </tr>
          <tr> 
            <td><a href="aspcheck.asp" target="main">ASP������̽��</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Admin_UploadFileManage.asp">�ϴ��ļ�����</a></td>
          </tr>
          <tr> 
            <td><a target="main" href="Help.asp">ϵͳ��<font color="000000">��</font></a></td>
          </tr>
        </table>
      </div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu2" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(2)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>��ҵ��Ϣ </span></td>
	</tr>
	<tr>
		<td id="submenu2" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a href="AdminAboutusAdd.asp" target="main">������ҵ��Ϣ</a></td>
          </tr>
          <tr> 
            <td><a href="AdminAboutus.asp" target="main">������ҵ��Ϣ</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_Culture.asp" target="main">������ҵ�Ļ�</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_CultureNewsAdd.asp" target="main">������ҵ�Ļ�</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu3" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(3)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>��Ʒ���� </span></td>
	</tr>
	<tr>
		<td id="submenu3" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="ClassManage.asp"><font color="000000">��Ʒ���</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="ProductManage.asp"><font color="000000">��Ʒ����</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="ProductAdd.asp"><font color="000000">��Ӳ�Ʒ</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="ProductCheck.asp"><font color="000000">��˲�Ʒ</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu4" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(4)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>ѯ�۹��� </span></td>
	</tr>
	<tr>
		<td id="submenu4" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a target="main" href="Admin_Order.asp"><font color="000000">ѯ�۹���</font></a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>		
    <td id="imgmenu14" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(14)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>�������� </span></td>
	</tr>
	<tr>
		<td id="submenu14" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a href="Down_add.asp" target="main">������س���</a></td>
          </tr>
          <tr> 
            <td><a href="Down_Manage.asp" target="main">�������س���</a></td>
          </tr>
          <tr> 
            <td><a href="Down_ClassManage.asp" target="main">���س������</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>


	<tr>
		
    <td id="imgmenu5" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(5)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>��Ա���� </span></td>
	</tr>
	<tr>
		<td id="submenu5" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="UserManage.asp"><font color="000000">ע���Ա����</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu6" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(6)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>���Ź���/FAQ </span></td>
	</tr>
	<tr>
		<td id="submenu6" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td> <a href="News_add.asp" target="main">�����������</a> </td>
          </tr>
          <tr> 
            <td><a href="News_Manage.asp" target="main">����ȫ������</a></td>
          </tr>
          <tr> 
            <td><a href="News_ClassManage.asp" target="main">�����������</a></td>
          </tr>
          <tr>
            <td><a href="Admin_faq.asp" target="main">��������</a></td>
          </tr>
          <tr>
            <td><a href="Admin_faqNewsAdd.asp" target="main">��ӳ�������</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>
	
	<tr>
		
    <td id="imgmenu7" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(7)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>���Թ��� </span></td>
	</tr>
	<tr>
		<td id="submenu7" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_Feedback.asp"><font color="000000">���Թ���</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_FeedbackAdd.asp"><font color="000000">����Ա����</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu8" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(8)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>�������� </span></td>
	</tr>
	<tr>
		<td id="submenu8" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_Honor.asp"><font color="000000">��ҵ��������</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HonorAdd.asp"><font color="000000">�����ҵ����</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_CompVisualize.asp"><font color="000000">��ҵ�������</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_CompVisualizeAdd.asp"><font color="000000">�����ҵ����</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu10" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(10)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>Ӫ������/�ͻ� </span></td>
	</tr>
	<tr>
		<td id="submenu10" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
		  <table cellspacing="3" cellpadding="0" width="130" align="center">
            <tr>
              <td><a target="main" href="market_input.asp"><font color="000000">���ӹ���Ӫ������</font></a></td>
            </tr>
            <tr>
              <td><a target="main" href="market.asp">�������Ӫ������</a></td>
            </tr>
            <tr>
              <td><a target="main" href="Admin_OverseasMarket.asp"><font color="000000">�����г�</font></a></td>
            </tr>
            <tr>
              <td><a href="class_customer.asp" target="main">����ͻ����</a></td>
            </tr>
            <tr>
              <td><a target="main" href="zp_input.asp">���ӿͻ�</a></td>
            </tr>
            <tr>
              <td><a target="main" href="zp.asp">����ͻ�</a></td>
            </tr>
          </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu9" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(9)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>�˲Ź��� </span></td>
	</tr>
	<tr>
		<td id="submenu9" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_HrDemand.asp"><font color="000000">��Ƹ����</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HrDemandAdd.asp"><font color="000000">������Ƹ</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HrManage.asp"><font color="000000">ӦƸ����</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="Admin_HrPolicy.asp"><font color="000000">�˲Ų���</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>
		
    <td id="imgmenu11" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(11)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>������� </span></td>
	</tr>
	<tr>
		<td id="submenu11" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="VoteManage.asp"><font color="000000">��������</font></a></td>
          </tr>
				<tr>
            <td><a target="main" href="VoteAdd.asp"><font color="000000">����µ���</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

	<tr>		
    <td id="imgmenu12" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(12)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>�ʼ��б� </span></td>
	</tr>
	<tr>
		<td id="submenu12" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a target="main" href="AdminMaillist.asp"><font color="000000">�ʼ��б�</font></a></td>
          </tr>
          <tr> 
            <td><a href="AdminMaillist.asp?Action=Export" target="main">�б���</a></td>
          </tr>
        </table>
		</div>
		<br>
		</td>
	</tr>

	<tr>		
    <td id="imgmenu13" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(13)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>�������� </span></td>
	</tr>
	<tr>
		<td id="submenu13" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px">
			<table cellspacing="3" cellpadding="0" width="130" align="center">
				<tr>
            <td><a target="main" href="Admin_FriendLinks.asp"><font color="000000">�������ӹ���</font></a></td>
          </tr>
			</table>
		</div>
		<br>
		</td>
	</tr>

<tr>		
    <td id="imgmenu15" class="menu_title" onMouseOver="this.className='menu_title2';" onClick="showsubmenu(15)" onMouseOut="this.className='menu_title';" style="cursor:hand" background="image/menudown.gif" height="25"> 
      <span>ģ����� </span></td>
  </tr>
	<tr>
		<td id="submenu15" style="DISPLAY: none">
		<div class="sec_menu" style="WIDTH: 158px"> 
        <table cellspacing="3" cellpadding="0" width="130" align="center">
          <tr> 
            <td><a href="Admin_Skin.asp" target="main">��ɫģ�����</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_Skin.asp?Action=Add" target="main">��ɫģ�����</a></td>
          </tr>
          <tr>
            <td><a href="Admin_Skin.asp?Action=Export" target="main">��ɫģ�嵼��</a></td>
          </tr>
          <tr> 
            <td><a href="Admin_Skin.asp?Action=Import" target="main">��ɫģ�嵼��</a></td>
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
      <span>��Ȩ����</span> </td>
	</tr>
	<tr>
		<td>
		<div class="sec_menu" style="WIDTH: 158px">
			<div align="center">
			<table cellspacing="4" cellpadding="3">
				<tr>
					
              <td width="141" height="20" style="line-height: 150%;">
                Copyright��<br><a target="_blank" href="/"><font color="000000"
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
{tt.doScroll(Dir);Timer=setTimeout('aa("'+Dir+'")',100)}//����100Ϊ�����ٶ�
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
top.document.title="����������ҵ��վ����ϵͳ��Ư������ҵ��վ����ϵͳ"; 
</script>
</BODY></HTML>
