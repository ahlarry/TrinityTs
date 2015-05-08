<%
'*************************************************
'函数名：gotTopic
'作  用：截字符串，汉字一个算两个字符，英文算一个字符
'参  数：str   ----原字符串
'       strlen ----截取长度
'返回值：截取后的字符串
'*************************************************
function gotTopic(str,strlen)
	if str="" then
		gotTopic=""
		exit function
	end if
	dim l,t,c, i
	str=replace(replace(replace(replace(str,"&nbsp;"," "),"&quot;",chr(34)),"&gt;",">"),"&lt;","<")
	l=len(str)
	t=0
	for i=1 to l
		c=Abs(Asc(Mid(str,i,1)))
		if c>255 then
			t=t+2
		else
			t=t+1
		end if
		if t>=strlen then
			gotTopic=left(str,i) & "…"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'***********************************************
'函数名：JoinChar
'作  用：向地址中加入 ? 或 &
'参  数：strUrl  ----网址
'返回值：加了 ? 或 & 的网址
'pos=InStr(1,"abcdefg","cd") 
'则pos会返回3表示查找到并且位置为第三个字符开始。
'这就是“查找”的实现，而“查找下一个”功能的
'实现就是把当前位置作为起始位置继续查找。
'***********************************************
function JoinChar(strUrl)
	if strUrl="" then
		JoinChar=""
		exit function
	end if
	if InStr(strUrl,"?")<len(strUrl) then 
		if InStr(strUrl,"?")>1 then
			if InStr(strUrl,"&")<len(strUrl) then 
				JoinChar=strUrl & "&"
			else
				JoinChar=strUrl
			end if
		else
			JoinChar=strUrl & "?"
		end if
	else
		JoinChar=strUrl
	end if
end function

'***********************************************
'过程名：showpage
'作  用：显示“上一页 下一页”等信息
'参  数：sfilename  ----链接地址
'       totalnumber ----总数量
'       maxperpage  ----每页数量
'       ShowTotal   ----是否显示总数量
'       ShowAllPages ---是否用下拉列表显示所有页面以供跳转。有某些页面不能使用，否则会出现JS错误。
'       strUnit     ----计数单位
'***********************************************
sub showpage(sfilename,totalnumber,maxperpage,ShowTotal,ShowAllPages,strUnit)
  
	dim n, i,strTemp,strUrl
	if totalnumber mod maxperpage=0 then
    	n= totalnumber \ maxperpage
  	else
    	n= totalnumber \ maxperpage+1
  	end if
  	strTemp= "<table align='right'><form name='showpages' method='Post' action='" & sfilename & "'><tr><td>"
	if ShowTotal=true then 
		strTemp=strTemp & "共 <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "首页 上一页&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>首页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>上一页</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "下一页 尾页"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>下一页</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>尾页</a>"
  	end if
   	strTemp=strTemp & "&nbsp;页次：<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>页 "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/页"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;转到：<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">第" & i & "页</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp	
end sub

'********************************************
'函数名：IsValidEmail
'作  用：检查Email地址合法性
'参  数：email ----要检查的Email地址
'返回值：True  ----Email地址合法
'       False ----Email地址不合法
'********************************************
function IsValidEmail(email)
	dim names, name, i, c
	IsValidEmail = true
	names = Split(email, "@")
	if UBound(names) <> 1 then
	   IsValidEmail = false
	   exit function
	end if
	for each name in names
		if Len(name) <= 0 then
			IsValidEmail = false
    		exit function
		end if
		for i = 1 to Len(name)
		    c = Lcase(Mid(name, i, 1))
			if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
		       IsValidEmail = false
		       exit function
		     end if
	   next
	   if Left(name, 1) = "." or Right(name, 1) = "." then
    	  IsValidEmail = false
	      exit function
	   end if
	next
	if InStr(names(1), ".") <= 0 then
		IsValidEmail = false
	   exit function
	end if
	i = Len(names(1)) - InStrRev(names(1), ".")
	if i <> 2 and i <> 3 then
	   IsValidEmail = false
	   exit function
	end if
	if InStr(email, "..") > 0 then
	   IsValidEmail = false
	end if
end function

'**************************************************
'函数名：strLength
'作  用：求字符串长度。汉字算两个字符，英文算一个字符。
'参  数：str  ----要求长度的字符串
'返回值：字符串长度
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("中国")=2)
	if WINNT_CHINESE then
        dim l,t,c
        dim i
        l=len(str)
        t=l
        for i=1 to l
        	c=asc(mid(str,i,1))
            if c<0 then c=c+65536
            if c>255 then
                t=t+1
            end if
        next
        strLength=t
    else 
        strLength=len(str)
    end if
    if err.number<>0 then err.clear
end function

'****************************************************
'函数名：SendMail
'作  用：用Jmail组件发送邮件
'参  数：ServerAddress  ----服务器地址
'        AddRecipient  ----收信人地址
'        Subject       ----主题
'        Body          ----信件内容
'        Sender        ----发信人地址
'****************************************************
function SendMail(MailServerAddress,AddRecipient,Subject,Body,Sender,MailFrom)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	if err then
		SendMail= "<br><li>没有安装JMail组件</li>"
		err.clear
		exit function
	end if
	JMail.Logging=True
	JMail.Charset="gb2312"
	JMail.ContentType = "text/html"
	JMail.ServerAddress=MailServerAddress
	JMail.AddRecipient=AddRecipient
	JMail.Subject=Subject
	JMail.Body=MailBody
	JMail.Sender=Sender
	JMail.From = MailFrom
	JMail.Priority=1
	JMail.Execute 
	Set JMail=nothing 
	if err then 
		SendMail=err.description
		err.clear
	else
		SendMail="OK"
	end if
end function

'****************************************************
'过程名：WriteErrMsg
'作  用：显示错误提示信息
'参  数：无
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>错误信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='20' class='title'><strong>错误信息</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>产生错误的可能原因：</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='title'><a href='" & Request.ServerVariables("HTTP_REFERER") & "'>【返回】</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'过程名：WriteSuccessMsg
'作  用：显示成功提示信息
'参  数：无
'****************************************************
sub WriteSuccessMsg(SuccessMsg,StrUrl)
	dim strSuccess
	If StrUrl="" Then StrUrl=Request.ServerVariables("HTTP_REFERER")
	strSuccess=strSuccess & "<html><head><title>成功信息</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td height='20' class='title'><strong>恭喜你！</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr><td height='100' class='tdbg' valign='top'><br>" & SuccessMsg &"<br></td></tr><tr><td class='tdbg' align='center'><br>" & AutoRefresh(3,StrUrl) & "<br></td></tr>"&vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td class='title'><a href='" & StrUrl & "'>【返回】</a></td></tr>" & vbcrlf
	strSuccess=strSuccess & "</table>" & vbcrlf
	strSuccess=strSuccess & "</body></html>" & vbcrlf
	response.write strSuccess
end sub

function getFileExtName(fileName)
    dim pos
    pos=instrrev(filename,".")
    if pos>0 then 
        getFileExtName=mid(fileName,pos+1)
    else
        getFileExtName=""
    end if
end function 

'==================================================
'过程名：ShowAnnounce
'作  用：显示本站公告信息
'        AnnounceNum  ----最多显示多少条公告
'==================================================
'sub ShowAnnounce(AnnounceNum)
'	dim sqlAnnounce,rsAnnounce,i
'	if AnnounceNum>0 and AnnounceNum<=10 then
'		sqlAnnounce="select top " & AnnounceNum
'	else
'		sqlAnnounce="select top 10"
'	end if
'	sqlAnnounce=sqlAnnounce & " * from affiche order by ID Desc"	
'	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
'	rsAnnounce.open sqlAnnounce,conn,1,1
'	if rsAnnounce.bof and rsAnnounce.eof then 
'		AnnounceCount=0
'		response.write "<p>&nbsp;&nbsp;没有公告</p>" 
'	else 
'		AnnounceCount=rsAnnounce.recordcount		
'			response.Write "本站公告："
'			do while not rsAnnounce.eof   
'				response.Write "&nbsp;<a href='#' onclick=""javascript:window.open('Affiche.asp?ID=" & rsAnnounce("id") &"', newwindow',' 'height=450, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'><font color='#FF0000'>" &rsAnnounce("title") & "</font></a>"
'				rsAnnounce.movenext
'				i=i+1				  
'			loop       		
'	end if  
'	rsAnnounce.close
'	set rsAnnounce=nothing
'end sub
sub ShowAnnounce(AnnounceNum) '显示最新标题及数量
	dim sqlAnnounce,rsAnnounce,i
	if AnnounceNum>0 and AnnounceNum<=10 then
		sqlAnnounce="select top " & AnnounceNum
	else
		sqlAnnounce="select top 10"
	end if
	sqlAnnounce=sqlAnnounce & " * from affiche order by ID Desc"	
	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
	rsAnnounce.open sqlAnnounce,conn,1,1
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>&nbsp;&nbsp;没有公告</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount		
			do while not rsAnnounce.eof   
				response.Write "&nbsp;&nbsp;<a href='#' onclick=""javascript:window.open('Affiche.asp?ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'><font color='#000000'>" &rsAnnounce("title") & "</font></a>"
				rsAnnounce.movenext
				i=i+1				  
			loop       		
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
end sub

sub ShowNewBBS() '显示最新一条内容
	dim sqlAnnounce,rsAnnounce,i
	sqlAnnounce=sqlAnnounce & "select top 4 * from affiche order by ID Desc"	
	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
	rsAnnounce.open sqlAnnounce,conn,1,1
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>&nbsp;&nbsp;没有公告</p>" 
	else 
		AnnounceCount=rsAnnounce.recordcount		
			do while not rsAnnounce.eof   
				response.Write "<a href='#' onclick=""javascript:window.open('Affiche.asp?ID=" & rsAnnounce("id") &"', 'newwindow', 'height=300, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')""><font color='#11455E'>" &rsAnnounce("Title")& "</font></a><br><br>"
				rsAnnounce.movenext
				i=i+1				  
			loop       		
	end if  
	rsAnnounce.close
	set rsAnnounce=nothing
end sub

'==================================================
'过程名：ShowFriendLinks
'作  用：显示友情链接站点
'参  数：LinkType  ----链接方式，1为LOGO链接，2为文字链接
'       SiteNum   ----最多显示多少个站点
'       Cols      ----分几列显示
'       ShowType  ----显示方式。1为向上滚动，2为横向列表，3为下拉列表框
'==================================================
sub ShowFriendLinks(LinkType,SiteNum,Cols,ShowType)
	dim sqlLink,rsLink,SiteCount,i,strLink
	if LinkType<>1 and LinkType<>2 then
		LinkType=1
	else
		LinkType=Cint(LinkType)
	end if
	if SiteNum<=0 or SiteNum>100 then
		SiteNum=10
	end if
	if Cols<=0 or Cols>20 then
		Cols=10
	end if
	if ShowType=1 then'
        strLink=strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"    '新增加的代码
	elseif ShowType=3 then
		strLink=strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>友情文字链接站点</option>"
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "<table width='100%' cellSpacing='1'><tr align='center' >"
	end if
	
	sqlLink="select top " & SiteNum & " * from FriendLinks where IsOK=True and LinkType=" & LinkType & " order by IsGood,id desc"
	set rsLink=server.createobject("adodb.recordset")
	rsLink.open sqlLink,conn,1,1
	if rsLink.bof and rsLink.eof then
		if ShowType=1 or ShowType=2 then
	  		for i=1 to SiteNum
				strLink=strLink & "<td>"			
				strLink=strLink & "</td>"
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' >"
				end if
			next
		end if
	else
		SiteCount=rsLink.recordcount
		for i=1 to SiteCount
			if ShowType=1 or ShowType=2 then
			  if LinkType=1 then
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td>"
			  else
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='网站名称：" & rsLink("SiteName") & vbcrlf & "网站地址：" & rsLink("SiteUrl") & vbcrlf & "网站简介：" & rsLink("SiteIntro") & "'>" & rsLink("SiteName") & "</a></td>"
			  end if
			  if i mod Cols=0 and i<SiteNum then
				strLink=strLink & "</tr><tr align='center' >"
			  end if
			else
				strLink=strLink & "<option value='" & rsLink("SiteUrl") & "'>" & rsLink("SiteName") & "</option>"
			end if
			rsLink.moveNext
		next
		if SiteCount<SiteNum and (ShowType=1 or ShowType=2) then
			for i=SiteCount+1 to SiteNum
				if LinkType=1 then
					strLink=strLink & "<td width='88'></td>"
				else
					strLink=strLink & "<td width='88'></td>"
				end if
				if i mod Cols=0 and i<SiteNum then
					strLink=strLink & "</tr><tr align='center' >"
				end if
			next
		end if
	end if
	if ShowType=1 or ShowType=2 then
		strLink=strLink & "</tr></table>"
	end if
	if ShowType=1 then
        strLink=strLink & "</div><div id=rolllink2></div></div>"   '新增代码
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	response.write strLink
	if ShowType=1 then call RollFriendLinks()    '新增代码
	rsLink.close
	set rsLink=nothing
end sub

'==================================================
'过程名：RollFriendLinks
'作  用：滚动显示友情链接站点
'参  数：无
'==================================================
sub RollFriendLinks()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //克隆rolllink1为rolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //当滚动至rolllink1与rolllink2交界时
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink跳到最顶端
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //设置定时器
   rolllink.onmouseover=function() {clearInterval(MyMar)}//鼠标移上时清除定时器达到滚动停止的目的
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//鼠标移开时重设定时器
</script>
<%
end sub
%>

<%
'--------------------------------------自动跳转页面代码---------------------------------------------
Function AutoRefresh(tTime,prePage)			'tTime为自动跳转的时间,prePage为跳转页面
	If prePage="" Then prePage=Request.ServerVariables("HTTP_REFERER")
	strCode="<span id=""downclock"" name=""downclock"">"&tTime&"</span> 秒后自动跳转到: "&_
		"<a href="""&prepage&""">前一页面</a>" &_
		vbcrlf & "<meta http-equiv=""refresh"" content="""& tTime &";url="&prepage&""">" &_
		vbcrlf & "<script language=""javascript"">" &_
		vbcrlf & "var totaltime = "&tTime&";	//倒计时秒数" &_
		vbcrlf & "function countDown()" &_
		vbcrlf & "{" &_
		vbcrlf & "downclock.innerHTML = totaltime;" &_
		vbcrlf & "window.setTimeout('countDown();',1000);"&_
		vbcrlf & "totaltime -= 1;"&_
		vbcrlf & "}" &_
		vbcrlf & "window.setTimeout('countDown();',1);"&_
		vbcrlf & "</script> "
	AutoRefresh=strCode
End Function
%>
<script language="javascript">
//==================================================
//TITLE及ALT提示效果
//==================================================
var pltsPop=null;
    var pltsoffsetX = 10; // 弹出窗口位于鼠标左侧或者右侧的距离；3-12 合适
    var pltsoffsetY = 15; // 弹出窗口位于鼠标下方的距离；3-12 合适
    var pltsTitle="";
    document.write('<div id=pltsTipLayer style="display: none;position: absolute; z-index:10001"></div>');
    function pltsinits()
    {
     document.onmouseover = plts;
     document.onmousemove = moveToMouseLoc;
    }
    function plts()
    { var o=event.srcElement;
     if(o.alt!=null && o.alt!=""){o.dypop=o.alt;o.alt=""};
     if(o.title!=null && o.title!=""){o.dypop=o.title;o.title=""};
     pltsPop=o.dypop;
     if(pltsPop!=null&&pltsPop!=""&&typeof(pltsPop)!="undefined")
     {
     pltsTipLayer.style.left=-1000;
     pltsTipLayer.style.display='';
     var Msg=pltsPop.replace(/\n/g,"<br>");
     Msg=Msg.replace(/\0x13/g,"<br>");
     var re=/\{(.[^\{]*)\}/ig;
     if(!re.test(Msg))pltsTitle="挤模厂调试部:";
     else{
     re=/\{(.[^\{]*)\}(.*)/ig;
     pltsTitle=Msg.replace(re,"$1")+" ";
     re=/\{(.[^\{]*)\}/ig;
     Msg=Msg.replace(re,"");
     Msg=Msg.replace("<br>","");}
     var content =
     '<table style="FILTER:alpha(opacity=90);border: 1px solid #cccccc" id="toolTipTalbe" cellspacing="1" cellpadding="0"><tr><td width="100%"><table bgcolor="#ffffff" cellspacing="0" cellpadding="0">'+
     '<tr id="pltsPoptop"><td height="20" bgcolor="#0094bb"><font color="#ffffff"><b><p id="topleft" align="left">↖'+pltsTitle+'</p><p id="topright" align="right" style="display:none">'+pltsTitle+'↗</font></b></font></td></tr>'+
     '<tr><td "+attr+" style="padding-left:10px;padding-right:10px;padding-top: 8px;padding-bottom:6px;line-height:140%">'+Msg+'</td></tr>'+
     '<tr id="pltsPopbot" style="display:none"><td height="20" bgcolor="#0094bb"><font color="#ffffff"><b><p id="botleft" align="left">↙'+pltsTitle+'</p><p id="botright" align="right" style="display:none">'+pltsTitle+'↘</font></b></font></td></tr>'+
     '</table></td></tr></table>';
     pltsTipLayer.innerHTML=content;
     toolTipTalbe.style.width=Math.min(pltsTipLayer.clientWidth,document.body.clientWidth/2.2);
     moveToMouseLoc();
     return true;
     }
     else
     {
     pltsTipLayer.innerHTML='';
     pltsTipLayer.style.display='none';
     return true;
     }
    } 

    function moveToMouseLoc()
    {
     if(pltsTipLayer.innerHTML=='')return true;
     var MouseX=event.x;
     var MouseY=event.y;
     var popHeight=pltsTipLayer.clientHeight;
     var popWidth=pltsTipLayer.clientWidth;
     if(MouseY+pltsoffsetY+popHeight>document.body.clientHeight)
     {
     popTopAdjust=-popHeight-pltsoffsetY*1.5;
     pltsPoptop.style.display="none";
     pltsPopbot.style.display="";
     }
     else
     {
     popTopAdjust=0;
     pltsPoptop.style.display="";
     pltsPopbot.style.display="none";
     }
     if(MouseX+pltsoffsetX+popWidth>document.body.clientWidth)
     {
     popLeftAdjust=-popWidth-pltsoffsetX*2;
     topleft.style.display="none";
     botleft.style.display="none";
     topright.style.display="";
     botright.style.display="";
     }
     else
     {
     popLeftAdjust=0;
     topleft.style.display="";
     botleft.style.display="";
     topright.style.display="none";
     botright.style.display="none";
     }
     pltsTipLayer.style.left=MouseX+pltsoffsetX+document.body.scrollLeft+popLeftAdjust;
     pltsTipLayer.style.top=MouseY+pltsoffsetY+document.body.scrollTop+popTopAdjust;
     return true;
    }
    pltsinits(); 
//==================================================
//又一款日历控件
//调用方法:
//日期包含时间：CalendarWebControl.show(this,true,this.value)
//日期不包含时间：CalendarWebControl.show(this,false,this.value)
//==================================================    
function atCalendarControl(){ 
var calendar=this; 
this.calendarPad=null; 
this.prevMonth=null; 
this.nextMonth=null; 
this.prevYear=null; 
this.nextYear=null; 
this.goToday=null; 
this.calendarClose=null; 
this.calendarAbout=null; 
this.head=null; 
this.body=null; 
this.today=[]; 
this.currentDate=[]; 
this.sltDate; 
this.target; 
this.source; 

/************** 加入日历底板及阴影 *********************/ 
this.addCalendarPad=function(){ 
document.write("<div id='divCalendarpad' style='position:absolute;top:100;left:0;width:255;height:187;display:none;'>"); 
document.write("<iframe frameborder=0 height=189 width=250></iframe>"); 
document.write("<div style='position:absolute;top:2;left:2;width:250;height:187;background-color:#336699;'></div>"); 
document.write("</div>"); 
calendar.calendarPad=document.all.divCalendarpad; 
} 
/************** 加入日历面板 *********************/ 
this.addCalendarBoard=function(){ 
var BOARD=this; 
var divBoard=document.createElement("div"); 
calendar.calendarPad.insertAdjacentElement("beforeEnd",divBoard); 
divBoard.style.cssText="position:absolute;top:0;left:0;width:250;height:187;border:0 outset;background-color:buttonface;"; 

var tbBoard=document.createElement("table"); 
divBoard.insertAdjacentElement("beforeEnd",tbBoard); 
tbBoard.style.cssText="position:absolute;top:2;left:2;width:248;height:10;font-size:9pt;"; 
tbBoard.cellPadding=0; 
tbBoard.cellSpacing=1; 

/************** 设置各功能按钮的功能 *********************/ 
/*********** Calendar About Button ***************/ 
trRow = tbBoard.insertRow(0); 
calendar.calendarAbout=calendar.insertTbCell(trRow,0,"-","center"); 
calendar.calendarAbout.title="帮助 快捷键:H"; 
calendar.calendarAbout.onclick=function(){calendar.about();} 
/*********** Calendar Head ***************/ 
tbCell=trRow.insertCell(1); 
tbCell.colSpan=5; 
tbCell.bgColor="#99CCFF"; 
tbCell.align="center"; 
tbCell.style.cssText = "cursor:default"; 
calendar.head=tbCell; 
/*********** Calendar Close Button ***************/ 
tbCell=trRow.insertCell(2); 
calendar.calendarClose = calendar.insertTbCell(trRow,2,"x","center"); 
calendar.calendarClose.title="关闭 快捷键:ESC或X"; 
calendar.calendarClose.onclick=function(){calendar.hide();} 

/*********** Calendar PrevYear Button ***************/ 
trRow = tbBoard.insertRow(1); 
calendar.prevYear = calendar.insertTbCell(trRow,0,"<<<","center"); 
calendar.prevYear.title="上一年 快捷键:↑"; 
calendar.prevYear.onmousedown=function(){ 
calendar.currentDate[0]--; 
calendar.show(calendar.target,calendar.returnTime,calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(calendar.currentDate[2]),calendar.source); 
} 
/*********** Calendar PrevMonth Button ***************/ 
calendar.prevMonth = calendar.insertTbCell(trRow,1,"<<","center"); 
calendar.prevMonth.title="上一月 快捷键:←"; 
calendar.prevMonth.onmousedown=function(){ 
calendar.currentDate[1]--; 
if(calendar.currentDate[1]==0){ 
calendar.currentDate[1]=12; 
calendar.currentDate[0]--; 
} 
calendar.show(calendar.target,calendar.returnTime,calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(calendar.currentDate[2]),calendar.source); 
} 
/*********** Calendar Today Button ***************/ 
calendar.goToday = calendar.insertTbCell(trRow,2,"今天","center",3); 
calendar.goToday.title="选择今天 快捷键:T"; 
calendar.goToday.onclick=function(){ 
if(calendar.returnTime) 
calendar.sltDate=calendar.today[0]+"-"+calendar.formatTime(calendar.today[1])+"-"+calendar.formatTime(calendar.today[2])+" "+calendar.formatTime(calendar.today[3])+":"+calendar.formatTime(calendar.today[4]) 
else 
calendar.sltDate=calendar.today[0]+"-"+calendar.formatTime(calendar.today[1])+"-"+calendar.formatTime(calendar.today[2]); 
calendar.target.value=calendar.sltDate; 
calendar.hide(); 
//calendar.show(calendar.target,calendar.today[0]+"-"+calendar.today[1]+"-"+calendar.today[2],calendar.source); 
} 
/*********** Calendar NextMonth Button ***************/ 
calendar.nextMonth = calendar.insertTbCell(trRow,3,">","center"); 
calendar.nextMonth.title="下一月 快捷键:→"; 
calendar.nextMonth.onmousedown=function(){ 
calendar.currentDate[1]++; 
if(calendar.currentDate[1]==13){ 
calendar.currentDate[1]=1; 
calendar.currentDate[0]++; 
} 
calendar.show(calendar.target,calendar.returnTime,calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(calendar.currentDate[2]),calendar.source); 
} 
/*********** Calendar NextYear Button ***************/ 
calendar.nextYear = calendar.insertTbCell(trRow,4,">>","center"); 
calendar.nextYear.title="下一年 快捷键:↓"; 
calendar.nextYear.onmousedown=function(){ 
calendar.currentDate[0]++; 
calendar.show(calendar.target,calendar.returnTime,calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(calendar.currentDate[2]),calendar.source); 

} 

trRow = tbBoard.insertRow(2); 
var cnDateName = new Array("日","一","二","三","四","五","六"); 
for (var i = 0; i < 7; i++) { 
tbCell=trRow.insertCell(i) 
tbCell.innerText=cnDateName[i]; 
tbCell.align="center"; 
tbCell.width=35; 
tbCell.style.cssText="cursor:default;border:1 solid #99CCCC;background-color:#99CCCC;"; 
} 

/*********** Calendar Body ***************/ 
trRow = tbBoard.insertRow(3); 
tbCell=trRow.insertCell(0); 
tbCell.colSpan=7; 
tbCell.height=97; 
tbCell.vAlign="top"; 
tbCell.bgColor="#F0F0F0"; 

var tbBody=document.createElement("table"); 
tbCell.insertAdjacentElement("beforeEnd",tbBody); 
tbBody.style.cssText="position:relative;top:0;left:0;width:245;height:103;font-size:9pt;" 
tbBody.cellPadding=0; 
tbBody.cellSpacing=1; 
calendar.body=tbBody; 

/*********** Time Body ***************/ 
trRow = tbBoard.insertRow(4); 
tbCell=trRow.insertCell(0); 
calendar.prevHours = calendar.insertTbCell(trRow,0,"-","center"); 
calendar.prevHours.title="小时调整 快捷键:Home"; 
calendar.prevHours.onmousedown=function(){ 
calendar.currentDate[3]--; 
if(calendar.currentDate[3]==-1) calendar.currentDate[3]=23; 
calendar.bottom.innerText=calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]); 
} 
tbCell=trRow.insertCell(1); 
calendar.nextHours = calendar.insertTbCell(trRow,1,"+","center"); 
calendar.nextHours.title="小时调整 快捷键:End"; 
calendar.nextHours.onmousedown=function(){ 
calendar.currentDate[3]++; 
if(calendar.currentDate[3]==24) calendar.currentDate[3]=0; 
calendar.bottom.innerText=calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]); 
} 
tbCell=trRow.insertCell(2); 
tbCell.colSpan=3; 
tbCell.bgColor="#99CCFF"; 
tbCell.align="center"; 
tbCell.style.cssText = "cursor:default"; 
calendar.bottom=tbCell; 
tbCell=trRow.insertCell(3); 
calendar.prevMinutes = calendar.insertTbCell(trRow,3,"-","center"); 
calendar.prevMinutes.title="分钟调整 快捷键:PageUp"; 
calendar.prevMinutes.onmousedown=function(){ 
calendar.currentDate[4]--; 
if(calendar.currentDate[4]==-1) calendar.currentDate[4]=59; 
calendar.bottom.innerText=calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]); 
} 
tbCell=trRow.insertCell(4); 
calendar.nextMinutes = calendar.insertTbCell(trRow,4,"+","center"); 
calendar.nextMinutes.title="分钟调整 快捷键:PageDown"; 
calendar.nextMinutes.onmousedown=function(){ 
calendar.currentDate[4]++; 
if(calendar.currentDate[4]==60) calendar.currentDate[4]=0; 
calendar.bottom.innerText=calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]); 
} 

} 

/************** 加入功能按钮公共样式 *********************/ 
this.insertTbCell=function(trRow,cellIndex,TXT,trAlign,tbColSpan){ 
var tbCell=trRow.insertCell(cellIndex); 
if(tbColSpan!=undefined) tbCell.colSpan=tbColSpan; 

var btnCell=document.createElement("button"); 
tbCell.insertAdjacentElement("beforeEnd",btnCell); 
btnCell.value=TXT; 
btnCell.style.cssText="width:100%;border:1 outset;background-color:buttonface;"; 
btnCell.onmouseover=function(){ 
btnCell.style.cssText="width:100%;border:1 outset;background-color:#F0F0F0;"; 

} 
btnCell.onmouseout=function(){ 
btnCell.style.cssText="width:100%;border:1 outset;background-color:buttonface;"; 
} 
// btnCell.onmousedown=function(){ 
// btnCell.style.cssText="width:100%;border:1 inset;background-color:#F0F0F0;"; 
// } 
btnCell.onmouseup=function(){ 
btnCell.style.cssText="width:100%;border:1 outset;background-color:#F0F0F0;"; 
} 
btnCell.onclick=function(){ 
btnCell.blur(); 
} 
return btnCell; 
} 

this.setDefaultDate=function(){ 
var dftDate=new Date(); 
calendar.today[0]=dftDate.getYear(); 
calendar.today[1]=dftDate.getMonth()+1; 
calendar.today[2]=dftDate.getDate(); 
calendar.today[3]=dftDate.getHours(); 
calendar.today[4]=dftDate.getMinutes(); 
} 

/****************** Show Calendar *********************/ 
this.show=function(targetObject,returnTime,defaultDate,sourceObject){ 
if(targetObject==undefined) { 
alert("未设置目标对象. \n方法: ATCALENDAR.show(obj 目标对象,boolean 是否返回时间,string 默认日期,obj 点击对象);\n\n目标对象:接受日期返回值的对象.\n默认日期:格式为\"yyyy-mm-dd\",缺省为当前日期.\n点击对象:点击这个对象弹出calendar,默认为目标对象.\n"); 
return false; 
} 
else calendar.target=targetObject; 

if(sourceObject==undefined) calendar.source=calendar.target; 
else calendar.source=sourceObject; 

if(returnTime) calendar.returnTime=true; 
else calendar.returnTime=false; 

var firstDay; 
var Cells=new Array(); 
if((defaultDate==undefined) || (defaultDate=="")){ 
var theDate=new Array(); 
calendar.head.innerText = calendar.today[0]+"-"+calendar.formatTime(calendar.today[1])+"-"+calendar.formatTime(calendar.today[2]); 
calendar.bottom.innerText = calendar.formatTime(calendar.today[3])+":"+calendar.formatTime(calendar.today[4]); 

theDate[0]=calendar.today[0]; theDate[1]=calendar.today[1]; theDate[2]=calendar.today[2]; 
theDate[3]=calendar.today[3]; theDate[4]=calendar.today[4]; 
} 
else{ 
var Datereg=/^\d{4}-\d{1,2}-\d{2}$/ 
var DateTimereg=/^(\d{1,4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{1,2})$/ 
if((!defaultDate.match(Datereg)) && (!defaultDate.match(DateTimereg))){ 
alert("默认日期(时间)的格式不正确！\t\n\n默认可接受格式为:\n1、yyyy-mm-dd \n2、yyyy-mm-dd hh:mm\n3、(空)"); 
calendar.setDefaultDate(); 
return; 
} 

if(defaultDate.match(Datereg)) defaultDate=defaultDate+" "+calendar.today[3]+":"+calendar.today[4]; 
var strDateTime=defaultDate.match(DateTimereg); 
var theDate=new Array(4) 
theDate[0]=strDateTime[1]; 
theDate[1]=strDateTime[2]; 
theDate[2]=strDateTime[3]; 
theDate[3]=strDateTime[4]; 
theDate[4]=strDateTime[5]; 
calendar.head.innerText = theDate[0]+"-"+calendar.formatTime(theDate[1])+"-"+calendar.formatTime(theDate[2]); 
calendar.bottom.innerText = calendar.formatTime(theDate[3])+":"+calendar.formatTime(theDate[4]); 
} 
calendar.currentDate[0]=theDate[0]; 
calendar.currentDate[1]=theDate[1]; 
calendar.currentDate[2]=theDate[2]; 
calendar.currentDate[3]=theDate[3]; 
calendar.currentDate[4]=theDate[4]; 

theFirstDay=calendar.getFirstDay(theDate[0],theDate[1]); 
theMonthLen=theFirstDay+calendar.getMonthLen(theDate[0],theDate[1]); 
//calendar.setEventKey(); 

calendar.calendarPad.style.display=""; 
var theRows = Math.ceil((theMonthLen)/7); 
//清除旧的日历; 
while (calendar.body.rows.length > 0) { 
calendar.body.deleteRow(0) 
} 
//建立新的日历; 
var n=0;day=0; 
for(i=0;i<theRows;i++){ 
theRow=calendar.body.insertRow(i); 
for(j=0;j<7;j++){ 
n++; 
if(n>theFirstDay && n<=theMonthLen){ 
day=n-theFirstDay; 
calendar.insertBodyCell(theRow,j,day); 
} 

else{ 
var theCell=theRow.insertCell(j); 
theCell.style.cssText="background-color:#F0F0F0;cursor:default;"; 
} 
} 
} 

//****************调整日历位置**************// 
var offsetPos=calendar.getAbsolutePos(calendar.source);//计算对象的位置; 
if((document.body.offsetHeight-(offsetPos.y+calendar.source.offsetHeight-document.body.scrollTop))<calendar.calendarPad.style.pixelHeight){ 
var calTop=offsetPos.y-calendar.calendarPad.style.pixelHeight; 
} 
else{ 
var calTop=offsetPos.y+calendar.source.offsetHeight; 
} 
if((document.body.offsetWidth-(offsetPos.x+calendar.source.offsetWidth-document.body.scrollLeft))>calendar.calendarPad.style.pixelWidth){ 
var calLeft=offsetPos.x; 
} 
else{ 
var calLeft=calendar.source.offsetLeft+calendar.source.offsetWidth; 
} 
//alert(offsetPos.x); 
calendar.calendarPad.style.pixelLeft=calLeft; 
calendar.calendarPad.style.pixelTop=calTop; 
} 
/****************** 计算对象的位置 *************************/ 
this.getAbsolutePos = function(el) { 
var r = { x: el.offsetLeft, y: el.offsetTop }; 
if (el.offsetParent) { 
var tmp = calendar.getAbsolutePos(el.offsetParent); 
r.x += tmp.x; 
r.y += tmp.y; 
} 
return r; 
}; 

//************* 插入日期单元格 **************/ 
this.insertBodyCell=function(theRow,j,day,targetObject){ 
var theCell=theRow.insertCell(j); 
if(j==0) var theBgColor="#FF9999"; 
else var theBgColor="#FFFFFF"; 
if(day==calendar.currentDate[2]) var theBgColor="#CCCCCC"; 
if(day==calendar.today[2]) var theBgColor="#99FFCC"; 
theCell.bgColor=theBgColor; 
theCell.innerText=day; 
theCell.align="center"; 
theCell.width=35; 
theCell.style.cssText="border:1 solid #CCCCCC;cursor:hand;"; 
theCell.onmouseover=function(){ 
theCell.bgColor="#FFFFCC"; 
theCell.style.cssText="border:1 outset;cursor:hand;"; 
} 
theCell.onmouseout=function(){ 
theCell.bgColor=theBgColor; 
theCell.style.cssText="border:1 solid #CCCCCC;cursor:hand;"; 
} 
theCell.onmousedown=function(){ 
theCell.bgColor="#FFFFCC"; 
theCell.style.cssText="border:1 inset;cursor:hand;"; 
} 
theCell.onclick=function(){ 
if(calendar.returnTime) 
calendar.sltDate=calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(day)+" "+calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]) 
else 
calendar.sltDate=calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(day); 
calendar.target.value=calendar.sltDate; 
calendar.hide(); 
} 
} 
/************** 取得月份的第一天为星期几 *********************/ 
this.getFirstDay=function(theYear, theMonth){ 
var firstDate = new Date(theYear,theMonth-1,1); 
return firstDate.getDay(); 
} 
/************** 取得月份共有几天 *********************/ 

this.getMonthLen=function(theYear, theMonth) { 
theMonth--; 
var oneDay = 1000 * 60 * 60 * 24; 
var thisMonth = new Date(theYear, theMonth, 1); 
var nextMonth = new Date(theYear, theMonth + 1, 1); 
var len = Math.ceil((nextMonth.getTime() - thisMonth.getTime())/oneDay); 
return len; 
} 
/************** 隐藏日历 *********************/ 
this.hide=function(){ 
//calendar.clearEventKey(); 
calendar.calendarPad.style.display="none"; 

} 
/************** 从这里开始 *********************/ 
this.setup=function(defaultDate){ 
calendar.addCalendarPad(); 
calendar.addCalendarBoard(); 
calendar.setDefaultDate(); 
} 
/************** 格式化时间 *********************/ 
this.formatTime = function(str) { 
str = ("00"+str); 
return str.substr(str.length-2); 
} 

/************** 关于AgetimeCalendar *********************/ 
this.about=function(){ 
var strAbout = "\nWeb 日历选择输入控件操作说明:\n\n"; 
strAbout+="-\t: 关于\n"; 
strAbout+="x\t: 隐藏\n"; 
strAbout+="<<\t: 上一年\n"; 
strAbout+="<\t: 上一月\n"; 

strAbout+="今日\t: 返回当天日期\n"; 
strAbout+=">\t: 下一月\n"; 
strAbout+="<<\t: 下一年\n"; 
strAbout+="\nWeb日历选择输入控件\tVer:v1.0\n\n\t2007.12.20\t\n"; 
alert(strAbout); 
} 

document.onkeydown=function(){ 
if(calendar.calendarPad.style.display=="none"){ 
window.event.returnValue= true; 
return true ; 
} 
switch(window.event.keyCode){ 
case 27 : calendar.hide(); break; //ESC 
case 37 : calendar.prevMonth.onmousedown(); break;//← 
case 38 : calendar.prevYear.onmousedown();break; //↑ 
case 39 : calendar.nextMonth.onmousedown(); break;//→ 
case 40 : calendar.nextYear.onmousedown(); break;//↓ 
case 84 : calendar.goToday.onclick(); break;//T 
case 88 : calendar.hide(); break; //X 
case 72 : calendar.about(); break; //H 
case 36 : calendar.prevHours.onmousedown(); break;//Home 
case 35 : calendar.nextHours.onmousedown(); break;//End 
case 33 : calendar.prevMinutes.onmousedown();break; //PageUp 
case 34 : calendar.nextMinutes.onmousedown(); break;//PageDown 
} 
window.event.keyCode = 0; 
window.event.returnValue= false; 
} 

calendar.setup(); 
} 

var CalendarWebControl = new atCalendarControl();     

/***************************************************************
数字的验证
****************************************************************/
function valNum(ev)
   {
    var e = ev.keyCode;
    //允许的有大、小键盘的数字、小数点、负号，左右键，backspace, delete, Ctl+C, Ctl+V
    if(e != 48 && e != 49 && e != 50 && e != 51 && e != 52 && e != 53 && e != 54 && e != 55 && e != 56 && e != 57 && e != 96 && e != 97 && e != 98 && e != 99 && e != 100 && e != 101 && e != 102 && e != 103 && e != 104 && e != 105 && e != 109 && e != 110 && e != 189 && e != 190 && e != 37 && e != 39 && e != 13 && e != 8 && e != 46)
       {
        if(ev.ctrlKey == false)
           {
            //不允许的就清空!
            ev.returnValue = "";
        }
        else
           {
            //验证剪贴板里的内容是否为数字!
            valClip(ev);
        }
    }
}
//验证剪贴板里的内容是否为数字!
function valClip(ev)
   {
    //查看剪贴板的内容!
    var content = clipboardData.getData("Text");
    if(content != null)
       {
        try
           {
            var test = parseInt(content);
            var str = "" + test;
            if(isNaN(test) == true)
               {
                //如果不是数字将内容清空!
                clipboardData.setData("Text","");
            }
            else
               {
                if(str != content)
                    clipboardData.setData("Text", str);
            }
        }
        catch(e)
           {
            //清空出现错误的提示!
            alert("粘贴出现错误!");
        }
    }
   }
</script>