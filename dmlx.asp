<style> 
.sml_menu {font-size: 9pt; color: white; cursor: hand; font-family: Tahoma;} 
.font3 {font-size: 10.5pt; color: 147e19; font-family: Courier New;} 
.menuitem {font-size: 10.5pt; color: white; cursor: default; font-family: Courier New;} 
</style>
<script language="javascript"> 
var intDelay=70; //���ò˵���ʾ�ٶȣ�Խ��Խ����������100Ϊ�� 
var intInterval=5; //ÿ�θ��ĵ�͸���ȣ����С��10 
function MenuClick() 
{ 
if (LayerMenu.style.display=="") 
{ 
LayerMenu.style.display="none"; //���˵���ʾ��ʱ�򵥻��͹رղ˵� 
} 
else{ 
LayerMenu.filters.alpha.opacity=0; 
LayerMenu.style.display=""; 
GradientShow(); //���� 
  } 
} 

function GradientShow() //ʵ�ֵ���ĺ��� 
{ 
LayerMenu.filters.alpha.opacity+=intInterval; 
if (LayerMenu.filters.alpha.opacity<100) setTimeout("GradientShow()",intDelay); 
} 

function GradientClose() //ʵ�ֵ����ĺ��� 
{ 
LayerMenu.filters.alpha.opacity-=intInterval; 
if (LayerMenu.filters.alpha.opacity>0) { 
  setTimeout("GradientClose()",intDelay); 
  } 
else { 
  LayerMenu.style.display="none"; //���������˵������Ҫ��������������� 
  } 
} 

function ChangeBG() //�ı�˵���ı�����ɫ�������������ɫֵ���Ը�Ϊ����Ҫ�� 
{ 
oEl=event.srcElement; 
if (oEl.style.background!="#03699A") { 
  oEl.style.background="#03699A"; 
  } 
  else { 
  oEl.style.background="#DE9E20"; 
  } 
} 

function ItemClick() //�ڲ˵����ϵ��������Ӧ���� 
{ 
oEl=event.srcElement; 
oLink=oEl.all.tags( "A" ); 
if( oLink.length ) 
{ 
oLink[0].click(); 
GradientClose(); 
} 
} 
</script>
<!--�˵��㿪ʼ-->
<div id=LayerMenu style="Position:Absolute;Left:250px;Top:300px;Display:none;filter:alpha(opacity=0);" oncontextmenu="return false" onMouseover="window.event.cancelBubble = true;">
  <table border=0 cellpadding=0 cellspacing=0 bgcolor="#5CB300">
    <tr>
      <td height=2 bgcolor="#5CB300" colspan=2></td>
    </tr>
    <tr>
      <td width=2 valign=bottom bgcolor="#5CB300"></td>
      <td><table border=0 width=450 cellpadding=0 cellspacing=0 onselectstart="return false;" onClick="ItemClick();" onMouseover="ChangeBG();" onMouseout="ChangeBG();">
          <%
    Dim c_dmdj, c_dmsm, n
    c_dmdj="" : c_dmsm=""
	'ȡ������ȼ�
	strSql="select * from dmlx order by dm"
	set rs=server.createobject("adodb.recordset")
	rs.open strSql,conn,1,1	
	do while not rs.eof
		if c_dmdj <> "" then 
			c_dmdj = c_dmdj & "|||" & rs("dm")
		else
			c_dmdj = rs("dm")
		end if
		if c_dmsm <> "" then 
			c_dmsm = c_dmsm & "|||" & rs("sm")
		else
			c_dmsm = rs("sm")
		end if		
		rs.movenext
	loop
	rs.close
	c_dmdj = split(c_dmdj, "|||")
	c_dmsm = split(c_dmsm, "|||")
	    
for n = 0 to ubound(c_dmdj)%>
          <tr>
            <td class=menuitem width=50 height=20 style="background: #03699A;"><%=c_dmdj(n)%></td>
            <td class=menuitem height=20 style="background: #03699A;"><%=c_dmsm(n)%></td>
          </tr>
          <tr>
            <td height="1" colspan="2" bgcolor="#D6DCE3"></td>
          </tr>
          <%next%>
        </table></td>
    </tr>
  </table>
</div>
<!--�˵������-->
