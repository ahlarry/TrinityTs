<%
'*************************************************
'��������gotTopic
'��  �ã����ַ���������һ���������ַ���Ӣ����һ���ַ�
'��  ����str   ----ԭ�ַ���
'       strlen ----��ȡ����
'����ֵ����ȡ����ַ���
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
			gotTopic=left(str,i) & "��"
			exit for
		else
			gotTopic=str
		end if
	next
	gotTopic=replace(replace(replace(replace(gotTopic," ","&nbsp;"),chr(34),"&quot;"),">","&gt;"),"<","&lt;")
end function

'***********************************************
'��������JoinChar
'��  �ã����ַ�м��� ? �� &
'��  ����strUrl  ----��ַ
'����ֵ������ ? �� & ����ַ
'pos=InStr(1,"abcdefg","cd") 
'��pos�᷵��3��ʾ���ҵ�����λ��Ϊ�������ַ���ʼ��
'����ǡ����ҡ���ʵ�֣�����������һ�������ܵ�
'ʵ�־��ǰѵ�ǰλ����Ϊ��ʼλ�ü������ҡ�
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
'��������showpage
'��  �ã���ʾ����һҳ ��һҳ������Ϣ
'��  ����sfilename  ----���ӵ�ַ
'       totalnumber ----������
'       maxperpage  ----ÿҳ����
'       ShowTotal   ----�Ƿ���ʾ������
'       ShowAllPages ---�Ƿ��������б���ʾ����ҳ���Թ���ת����ĳЩҳ�治��ʹ�ã���������JS����
'       strUnit     ----������λ
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
		strTemp=strTemp & "�� <b>" & totalnumber & "</b> " & strUnit & "&nbsp;&nbsp;"
	end if
	strUrl=JoinChar(sfilename)
  	if CurrentPage<2 then
    		strTemp=strTemp & "��ҳ ��һҳ&nbsp;"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=1'>��ҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage-1) & "'>��һҳ</a>&nbsp;"
  	end if

  	if n-currentpage<1 then
    		strTemp=strTemp & "��һҳ βҳ"
  	else
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & (CurrentPage+1) & "'>��һҳ</a>&nbsp;"
    		strTemp=strTemp & "<a href='" & strUrl & "page=" & n & "'>βҳ</a>"
  	end if
   	strTemp=strTemp & "&nbsp;ҳ�Σ�<strong><font color=red>" & CurrentPage & "</font>/" & n & "</strong>ҳ "
    strTemp=strTemp & "&nbsp;<b>" & maxperpage & "</b>" & strUnit & "/ҳ"
	if ShowAllPages=True then
		strTemp=strTemp & "&nbsp;ת����<select name='page' size='1' onchange='javascript:submit()'>"   
    	for i = 1 to n   
    		strTemp=strTemp & "<option value='" & i & "'"
			if cint(CurrentPage)=cint(i) then strTemp=strTemp & " selected "
			strTemp=strTemp & ">��" & i & "ҳ</option>"   
	    next
		strTemp=strTemp & "</select>"
	end if
	strTemp=strTemp & "</td></tr></form></table>"
	response.write strTemp	
end sub

'********************************************
'��������IsValidEmail
'��  �ã����Email��ַ�Ϸ���
'��  ����email ----Ҫ����Email��ַ
'����ֵ��True  ----Email��ַ�Ϸ�
'       False ----Email��ַ���Ϸ�
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
'��������strLength
'��  �ã����ַ������ȡ������������ַ���Ӣ����һ���ַ���
'��  ����str  ----Ҫ�󳤶ȵ��ַ���
'����ֵ���ַ�������
'**************************************************
function strLength(str)
	ON ERROR RESUME NEXT
	dim WINNT_CHINESE
	WINNT_CHINESE    = (len("�й�")=2)
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
'��������SendMail
'��  �ã���Jmail��������ʼ�
'��  ����ServerAddress  ----��������ַ
'        AddRecipient  ----�����˵�ַ
'        Subject       ----����
'        Body          ----�ż�����
'        Sender        ----�����˵�ַ
'****************************************************
function SendMail(MailServerAddress,AddRecipient,Subject,Body,Sender,MailFrom)
	on error resume next
	Dim JMail
	Set JMail=Server.CreateObject("JMail.SMTPMail")
	if err then
		SendMail= "<br><li>û�а�װJMail���</li>"
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
'��������WriteErrMsg
'��  �ã���ʾ������ʾ��Ϣ
'��  ������
'****************************************************
sub WriteErrMsg()
	dim strErr
	strErr=strErr & "<html><head><title>������Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strErr=strErr & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strErr=strErr & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td height='20' class='title'><strong>������Ϣ</strong></td></tr>" & vbcrlf
	strErr=strErr & "  <tr><td height='100' class='tdbg' valign='top'><b>��������Ŀ���ԭ��</b><br>" & errmsg &"</td></tr>" & vbcrlf
	strErr=strErr & "  <tr align='center'><td class='title'><a href='" & Request.ServerVariables("HTTP_REFERER") & "'>�����ء�</a></td></tr>" & vbcrlf
	strErr=strErr & "</table>" & vbcrlf
	strErr=strErr & "</body></html>" & vbcrlf
	response.write strErr
end sub

'****************************************************
'��������WriteSuccessMsg
'��  �ã���ʾ�ɹ���ʾ��Ϣ
'��  ������
'****************************************************
sub WriteSuccessMsg(SuccessMsg,StrUrl)
	dim strSuccess
	If StrUrl="" Then StrUrl=Request.ServerVariables("HTTP_REFERER")
	strSuccess=strSuccess & "<html><head><title>�ɹ���Ϣ</title><meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & vbcrlf
	strSuccess=strSuccess & "<link href='style.css' rel='stylesheet' type='text/css'></head><body>" & vbcrlf
	strSuccess=strSuccess & "<table cellpadding=2 cellspacing=2 border=0 width=400 class='border' align=center>" & vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td height='20' class='title'><strong>��ϲ�㣡</strong></td></tr>" & vbcrlf
	strSuccess=strSuccess & "  <tr><td height='100' class='tdbg' valign='top'><br>" & SuccessMsg &"<br></td></tr><tr><td class='tdbg' align='center'><br>" & AutoRefresh(3,StrUrl) & "<br></td></tr>"&vbcrlf
	strSuccess=strSuccess & "  <tr align='center'><td class='title'><a href='" & StrUrl & "'>�����ء�</a></td></tr>" & vbcrlf
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
'��������ShowAnnounce
'��  �ã���ʾ��վ������Ϣ
'        AnnounceNum  ----�����ʾ����������
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
'		response.write "<p>&nbsp;&nbsp;û�й���</p>" 
'	else 
'		AnnounceCount=rsAnnounce.recordcount		
'			response.Write "��վ���棺"
'			do while not rsAnnounce.eof   
'				response.Write "&nbsp;<a href='#' onclick=""javascript:window.open('Affiche.asp?ID=" & rsAnnounce("id") &"', newwindow',' 'height=450, width=400, toolbar=no, menubar=no, scrollbars=auto, resizable=no, location=no, status=no')"" title='" & rsAnnounce("Content") & "'><font color='#FF0000'>" &rsAnnounce("title") & "</font></a>"
'				rsAnnounce.movenext
'				i=i+1				  
'			loop       		
'	end if  
'	rsAnnounce.close
'	set rsAnnounce=nothing
'end sub
sub ShowAnnounce(AnnounceNum) '��ʾ���±��⼰����
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
		response.write "<p>&nbsp;&nbsp;û�й���</p>" 
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

sub ShowNewBBS() '��ʾ����һ������
	dim sqlAnnounce,rsAnnounce,i
	sqlAnnounce=sqlAnnounce & "select top 4 * from affiche order by ID Desc"	
	Set rsAnnounce= Server.CreateObject("ADODB.Recordset")
	rsAnnounce.open sqlAnnounce,conn,1,1
	if rsAnnounce.bof and rsAnnounce.eof then 
		AnnounceCount=0
		response.write "<p>&nbsp;&nbsp;û�й���</p>" 
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
'��������ShowFriendLinks
'��  �ã���ʾ��������վ��
'��  ����LinkType  ----���ӷ�ʽ��1ΪLOGO���ӣ�2Ϊ��������
'       SiteNum   ----�����ʾ���ٸ�վ��
'       Cols      ----�ּ�����ʾ
'       ShowType  ----��ʾ��ʽ��1Ϊ���Ϲ�����2Ϊ�����б�3Ϊ�����б��
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
        strLink=strLink & "<div id=rolllink style=overflow:hidden;height:100;width:100><div id=rolllink1>"    '�����ӵĴ���
	elseif ShowType=3 then
		strLink=strLink & "<select name='FriendSite' onchange=""if(this.options[this.selectedIndex].value!=''){window.open(this.options[this.selectedIndex].value,'_blank');}""><option value=''>������������վ��</option>"
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
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>"
				if rsLink("LogoUrl")="" or rsLink("LogoUrl")="http://" then
					strLink=strLink & "<img src='images/nologo.gif' width='88' height='31' border='0'>"
				else
					strLink=strLink & "<img src='" & rsLink("LogoUrl") & "' width='88' height='31' border='0'>"
				end if
				strLink=strLink & "</a></td>"
			  else
				strLink=strLink & "<td width='88'><a href='" & rsLink("SiteUrl") & "' target='_blank' title='��վ���ƣ�" & rsLink("SiteName") & vbcrlf & "��վ��ַ��" & rsLink("SiteUrl") & vbcrlf & "��վ��飺" & rsLink("SiteIntro") & "'>" & rsLink("SiteName") & "</a></td>"
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
        strLink=strLink & "</div><div id=rolllink2></div></div>"   '��������
	elseif ShowType=3 then
		strLink=strLink & "</select>"
	end if
	response.write strLink
	if ShowType=1 then call RollFriendLinks()    '��������
	rsLink.close
	set rsLink=nothing
end sub

'==================================================
'��������RollFriendLinks
'��  �ã�������ʾ��������վ��
'��  ������
'==================================================
sub RollFriendLinks()
%>
<script>
   var rollspeed=30
   rolllink2.innerHTML=rolllink1.innerHTML //��¡rolllink1Ϊrolllink2
   function Marquee(){
   if(rolllink2.offsetTop-rolllink.scrollTop<=0) //��������rolllink1��rolllink2����ʱ
   rolllink.scrollTop-=rolllink1.offsetHeight  //rolllink�������
   else{
   rolllink.scrollTop++
   }
   }
   var MyMar=setInterval(Marquee,rollspeed) //���ö�ʱ��
   rolllink.onmouseover=function() {clearInterval(MyMar)}//�������ʱ�����ʱ���ﵽ����ֹͣ��Ŀ��
   rolllink.onmouseout=function() {MyMar=setInterval(Marquee,rollspeed)}//����ƿ�ʱ���趨ʱ��
</script>
<%
end sub
%>

<%
'--------------------------------------�Զ���תҳ�����---------------------------------------------
Function AutoRefresh(tTime,prePage)			'tTimeΪ�Զ���ת��ʱ��,prePageΪ��תҳ��
	If prePage="" Then prePage=Request.ServerVariables("HTTP_REFERER")
	strCode="<span id=""downclock"" name=""downclock"">"&tTime&"</span> ����Զ���ת��: "&_
		"<a href="""&prepage&""">ǰһҳ��</a>" &_
		vbcrlf & "<meta http-equiv=""refresh"" content="""& tTime &";url="&prepage&""">" &_
		vbcrlf & "<script language=""javascript"">" &_
		vbcrlf & "var totaltime = "&tTime&";	//����ʱ����" &_
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
//TITLE��ALT��ʾЧ��
//==================================================
var pltsPop=null;
    var pltsoffsetX = 10; // ��������λ������������Ҳ�ľ��룻3-12 ����
    var pltsoffsetY = 15; // ��������λ������·��ľ��룻3-12 ����
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
     if(!re.test(Msg))pltsTitle="��ģ�����Բ�:";
     else{
     re=/\{(.[^\{]*)\}(.*)/ig;
     pltsTitle=Msg.replace(re,"$1")+" ";
     re=/\{(.[^\{]*)\}/ig;
     Msg=Msg.replace(re,"");
     Msg=Msg.replace("<br>","");}
     var content =
     '<table style="FILTER:alpha(opacity=90);border: 1px solid #cccccc" id="toolTipTalbe" cellspacing="1" cellpadding="0"><tr><td width="100%"><table bgcolor="#ffffff" cellspacing="0" cellpadding="0">'+
     '<tr id="pltsPoptop"><td height="20" bgcolor="#0094bb"><font color="#ffffff"><b><p id="topleft" align="left">�I'+pltsTitle+'</p><p id="topright" align="right" style="display:none">'+pltsTitle+'�J</font></b></font></td></tr>'+
     '<tr><td "+attr+" style="padding-left:10px;padding-right:10px;padding-top: 8px;padding-bottom:6px;line-height:140%">'+Msg+'</td></tr>'+
     '<tr id="pltsPopbot" style="display:none"><td height="20" bgcolor="#0094bb"><font color="#ffffff"><b><p id="botleft" align="left">�L'+pltsTitle+'</p><p id="botright" align="right" style="display:none">'+pltsTitle+'�K</font></b></font></td></tr>'+
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
//��һ�������ؼ�
//���÷���:
//���ڰ���ʱ�䣺CalendarWebControl.show(this,true,this.value)
//���ڲ�����ʱ�䣺CalendarWebControl.show(this,false,this.value)
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

/************** ���������װ弰��Ӱ *********************/ 
this.addCalendarPad=function(){ 
document.write("<div id='divCalendarpad' style='position:absolute;top:100;left:0;width:255;height:187;display:none;'>"); 
document.write("<iframe frameborder=0 height=189 width=250></iframe>"); 
document.write("<div style='position:absolute;top:2;left:2;width:250;height:187;background-color:#336699;'></div>"); 
document.write("</div>"); 
calendar.calendarPad=document.all.divCalendarpad; 
} 
/************** ����������� *********************/ 
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

/************** ���ø����ܰ�ť�Ĺ��� *********************/ 
/*********** Calendar About Button ***************/ 
trRow = tbBoard.insertRow(0); 
calendar.calendarAbout=calendar.insertTbCell(trRow,0,"-","center"); 
calendar.calendarAbout.title="���� ��ݼ�:H"; 
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
calendar.calendarClose.title="�ر� ��ݼ�:ESC��X"; 
calendar.calendarClose.onclick=function(){calendar.hide();} 

/*********** Calendar PrevYear Button ***************/ 
trRow = tbBoard.insertRow(1); 
calendar.prevYear = calendar.insertTbCell(trRow,0,"<<<","center"); 
calendar.prevYear.title="��һ�� ��ݼ�:��"; 
calendar.prevYear.onmousedown=function(){ 
calendar.currentDate[0]--; 
calendar.show(calendar.target,calendar.returnTime,calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(calendar.currentDate[2]),calendar.source); 
} 
/*********** Calendar PrevMonth Button ***************/ 
calendar.prevMonth = calendar.insertTbCell(trRow,1,"<<","center"); 
calendar.prevMonth.title="��һ�� ��ݼ�:��"; 
calendar.prevMonth.onmousedown=function(){ 
calendar.currentDate[1]--; 
if(calendar.currentDate[1]==0){ 
calendar.currentDate[1]=12; 
calendar.currentDate[0]--; 
} 
calendar.show(calendar.target,calendar.returnTime,calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(calendar.currentDate[2]),calendar.source); 
} 
/*********** Calendar Today Button ***************/ 
calendar.goToday = calendar.insertTbCell(trRow,2,"����","center",3); 
calendar.goToday.title="ѡ����� ��ݼ�:T"; 
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
calendar.nextMonth.title="��һ�� ��ݼ�:��"; 
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
calendar.nextYear.title="��һ�� ��ݼ�:��"; 
calendar.nextYear.onmousedown=function(){ 
calendar.currentDate[0]++; 
calendar.show(calendar.target,calendar.returnTime,calendar.currentDate[0]+"-"+calendar.formatTime(calendar.currentDate[1])+"-"+calendar.formatTime(calendar.currentDate[2]),calendar.source); 

} 

trRow = tbBoard.insertRow(2); 
var cnDateName = new Array("��","һ","��","��","��","��","��"); 
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
calendar.prevHours.title="Сʱ���� ��ݼ�:Home"; 
calendar.prevHours.onmousedown=function(){ 
calendar.currentDate[3]--; 
if(calendar.currentDate[3]==-1) calendar.currentDate[3]=23; 
calendar.bottom.innerText=calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]); 
} 
tbCell=trRow.insertCell(1); 
calendar.nextHours = calendar.insertTbCell(trRow,1,"+","center"); 
calendar.nextHours.title="Сʱ���� ��ݼ�:End"; 
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
calendar.prevMinutes.title="���ӵ��� ��ݼ�:PageUp"; 
calendar.prevMinutes.onmousedown=function(){ 
calendar.currentDate[4]--; 
if(calendar.currentDate[4]==-1) calendar.currentDate[4]=59; 
calendar.bottom.innerText=calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]); 
} 
tbCell=trRow.insertCell(4); 
calendar.nextMinutes = calendar.insertTbCell(trRow,4,"+","center"); 
calendar.nextMinutes.title="���ӵ��� ��ݼ�:PageDown"; 
calendar.nextMinutes.onmousedown=function(){ 
calendar.currentDate[4]++; 
if(calendar.currentDate[4]==60) calendar.currentDate[4]=0; 
calendar.bottom.innerText=calendar.formatTime(calendar.currentDate[3])+":"+calendar.formatTime(calendar.currentDate[4]); 
} 

} 

/************** ���빦�ܰ�ť������ʽ *********************/ 
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
alert("δ����Ŀ�����. \n����: ATCALENDAR.show(obj Ŀ�����,boolean �Ƿ񷵻�ʱ��,string Ĭ������,obj �������);\n\nĿ�����:�������ڷ���ֵ�Ķ���.\nĬ������:��ʽΪ\"yyyy-mm-dd\",ȱʡΪ��ǰ����.\n�������:���������󵯳�calendar,Ĭ��ΪĿ�����.\n"); 
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
alert("Ĭ������(ʱ��)�ĸ�ʽ����ȷ��\t\n\nĬ�Ͽɽ��ܸ�ʽΪ:\n1��yyyy-mm-dd \n2��yyyy-mm-dd hh:mm\n3��(��)"); 
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
//����ɵ�����; 
while (calendar.body.rows.length > 0) { 
calendar.body.deleteRow(0) 
} 
//�����µ�����; 
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

//****************��������λ��**************// 
var offsetPos=calendar.getAbsolutePos(calendar.source);//��������λ��; 
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
/****************** ��������λ�� *************************/ 
this.getAbsolutePos = function(el) { 
var r = { x: el.offsetLeft, y: el.offsetTop }; 
if (el.offsetParent) { 
var tmp = calendar.getAbsolutePos(el.offsetParent); 
r.x += tmp.x; 
r.y += tmp.y; 
} 
return r; 
}; 

//************* �������ڵ�Ԫ�� **************/ 
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
/************** ȡ���·ݵĵ�һ��Ϊ���ڼ� *********************/ 
this.getFirstDay=function(theYear, theMonth){ 
var firstDate = new Date(theYear,theMonth-1,1); 
return firstDate.getDay(); 
} 
/************** ȡ���·ݹ��м��� *********************/ 

this.getMonthLen=function(theYear, theMonth) { 
theMonth--; 
var oneDay = 1000 * 60 * 60 * 24; 
var thisMonth = new Date(theYear, theMonth, 1); 
var nextMonth = new Date(theYear, theMonth + 1, 1); 
var len = Math.ceil((nextMonth.getTime() - thisMonth.getTime())/oneDay); 
return len; 
} 
/************** �������� *********************/ 
this.hide=function(){ 
//calendar.clearEventKey(); 
calendar.calendarPad.style.display="none"; 

} 
/************** �����￪ʼ *********************/ 
this.setup=function(defaultDate){ 
calendar.addCalendarPad(); 
calendar.addCalendarBoard(); 
calendar.setDefaultDate(); 
} 
/************** ��ʽ��ʱ�� *********************/ 
this.formatTime = function(str) { 
str = ("00"+str); 
return str.substr(str.length-2); 
} 

/************** ����AgetimeCalendar *********************/ 
this.about=function(){ 
var strAbout = "\nWeb ����ѡ������ؼ�����˵��:\n\n"; 
strAbout+="-\t: ����\n"; 
strAbout+="x\t: ����\n"; 
strAbout+="<<\t: ��һ��\n"; 
strAbout+="<\t: ��һ��\n"; 

strAbout+="����\t: ���ص�������\n"; 
strAbout+=">\t: ��һ��\n"; 
strAbout+="<<\t: ��һ��\n"; 
strAbout+="\nWeb����ѡ������ؼ�\tVer:v1.0\n\n\t2007.12.20\t\n"; 
alert(strAbout); 
} 

document.onkeydown=function(){ 
if(calendar.calendarPad.style.display=="none"){ 
window.event.returnValue= true; 
return true ; 
} 
switch(window.event.keyCode){ 
case 27 : calendar.hide(); break; //ESC 
case 37 : calendar.prevMonth.onmousedown(); break;//�� 
case 38 : calendar.prevYear.onmousedown();break; //�� 
case 39 : calendar.nextMonth.onmousedown(); break;//�� 
case 40 : calendar.nextYear.onmousedown(); break;//�� 
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
���ֵ���֤
****************************************************************/
function valNum(ev)
   {
    var e = ev.keyCode;
    //������д�С���̵����֡�С���㡢���ţ����Ҽ���backspace, delete, Ctl+C, Ctl+V
    if(e != 48 && e != 49 && e != 50 && e != 51 && e != 52 && e != 53 && e != 54 && e != 55 && e != 56 && e != 57 && e != 96 && e != 97 && e != 98 && e != 99 && e != 100 && e != 101 && e != 102 && e != 103 && e != 104 && e != 105 && e != 109 && e != 110 && e != 189 && e != 190 && e != 37 && e != 39 && e != 13 && e != 8 && e != 46)
       {
        if(ev.ctrlKey == false)
           {
            //������ľ����!
            ev.returnValue = "";
        }
        else
           {
            //��֤��������������Ƿ�Ϊ����!
            valClip(ev);
        }
    }
}
//��֤��������������Ƿ�Ϊ����!
function valClip(ev)
   {
    //�鿴�����������!
    var content = clipboardData.getData("Text");
    if(content != null)
       {
        try
           {
            var test = parseInt(content);
            var str = "" + test;
            if(isNaN(test) == true)
               {
                //����������ֽ��������!
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
            //��ճ��ִ������ʾ!
            alert("ճ�����ִ���!");
        }
    }
   }
</script>