<%@language=VBScript codepage=936 %>
<!--#include file="conn.asp"-->
<!--#include file="admin.asp"-->
<!-- #include file="Inc/Head.asp" -->
<!--#include file="../inc/config.asp"-->
<!--#include file="inc/function.asp"-->
<%
dim ObjInstalled,Action,FoundErr,ErrMsg
ObjInstalled=IsObjInstalled("Scripting.FileSystemObject")
Action=trim(request("Action"))
if Action="" then
	Action="ShowInfo"
end if
%>
<html>
<head>
<title>��վ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0"> 
<%
if Action="SaveConfig" then
	call SaveConfig()
	Response.Redirect "SiteConfig.asp" 
else
	call ShowConfig()
end if
if FoundErr=True then
	call WriteErrMsg()
end if
call CloseConn()

sub ShowConfig()
%>
<form method="POST" action="SiteConfig.asp" id="form" name="form">
  <table width="650" border="0" align="center" cellpadding="2" cellspacing="1" bgcolor="#000000" >
    <tr> 
      <td height="24" colspan="4" class="back_southidc"> <div align="center"><strong>�� 
          վ �� ��</strong></div></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"> <strong>��վ��Ϣ����</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ���ƣ�</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteName" type="text" id="SiteName" value="<%=SiteName%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ���⣺</strong></td>
      <td colspan="3" class="tdbg"> <input name="SiteTitle" type="text" id="SiteTitle" value="<%=SiteTitle%>" size="40" maxlength="50"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��վ��ַ��</strong><br>
        ����д����URL��ַ(������վͳ��)</td>
      <td colspan="3" class="tdbg"> <input name="SiteUrl" type="text" id="SiteUrl" value="<%=SiteUrl%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��ҵ�ʾ֣�</strong><br>
        ����д����URL��ַ</td>
      <td colspan="3" class="tdbg"> <input name="EnterpriseMail" type="text" id="EnterpriseMail" value="<%=EnterpriseMail%>" size="40" maxlength="255"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>LOGO��ַ��</strong><br> </td>
      <td colspan="3" class="tdbg"> <input name="LogoUrl" type="text" id="LogoUrl" value="<%=LogoUrl%>" size="40" maxlength="255"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" height="20" class="tdbg"><strong>��ҳBanner��ַ(ֻ����FLASH)��</strong><br> 
      </td>
      <td width="186" class="tdbg"> <input name="BannerUrl" type="text" id="BannerUrl" value="<%=BannerUrl%>" size="25" maxlength="255"> 
      </td>
      <td width="42" class="tdbg">�߶ȣ�</td>
      <td width="127" class="tdbg"> <input name="High" type="text" id="High" value="<%=High%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>վ��������</strong></td>
      <td colspan="3" class="tdbg"> <input name="WebmasterName" type="text" id="WebmasterName" value="<%=WebmasterName%>" size="40" maxlength="20"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>վ�����䣺</strong></td>
      <td colspan="3" class="tdbg"> <input name="WebmasterEmail" type="text" id="WebmasterEmail" value="<%=WebmasterEmail%>" size="40" maxlength="100"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>��Ȩ��Ϣ��</strong><br>
        ֧��HTML��ǣ�����ʹ��˫����</td>
      <td colspan="3" class="tdbg"> <textarea name="Copyright" cols="50" rows="8" id="Copyright"><%=Copyright%></textarea> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"><strong>��վѡ������</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��ҳÿҳ��Ʒ����������</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxPerPage_Default" type="text" id="MaxPerPage_Default" value="<%=MaxPerPage_Default%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��ҳ������Ѷ������</strong></td>
      <td colspan="3" class="tdbg"> <input name="New_count" type="text" id="New_count" value="<%=New_count%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��ҳ��Ʒ�б�����</strong></td>
      <td colspan="3" class="tdbg"> <input name="Product_count" type="text" id="Product_count" value="<%=Product_count%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>��Ʒ����ҳÿҳ��������</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxPerPage_Search" type="text" id="MaxPerPage_Search" value="<%=MaxPerPage_Search%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>ÿҳ��ʾ��Լ�ַ�����</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxPerPage_Content" type="text" id="MaxPerPage_Content" value="<%=MaxPerPage_Content%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td class="tdbg"><strong>�Ƿ������������۹��ܣ�</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="NewsComment" value="Yes" <%if NewsComment="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="NewsComment" value="No" <%if NewsComment="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�Ƿ����ò�Ʒ��˹��ܣ�</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="EnableArticleCheck" value="Yes" <%if EnableArticleCheck="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="EnableArticleCheck" value="No" <%if EnableArticleCheck="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�Ƿ񿪷��ļ��ϴ���</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="EnableUploadFile" value="Yes" <%if EnableUploadFile="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="EnableUploadFile" value="No" <%if EnableUploadFile="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�Ƿ���ʾ���棺</strong></td>
      <td colspan="3" class="tdbg"> <input type="radio" name="PopAnnounce" value="Yes" <%if PopAnnounce="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="PopAnnounce" value="No" <%if PopAnnounce="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�������ŵ������</strong></td>
      <td colspan="3" class="tdbg"> <input name="HitsOfHot" type="text" id="HitsOfHot" value="<%=HitsOfHot%>" size="6" maxlength="5"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�ϴ��ļ���С���ƣ�</strong><br>
        ���鲻Ҫ����1024K������Ӱ�����������<strong>��</strong></td>
      <td colspan="3" class="tdbg"> <input name="MaxFileSize" type="text" id="MaxFileSize" value="<%=MaxFileSize%>" size="6" maxlength="5">
        K</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>����ϴ��ļ���Ŀ¼��</strong><br>
        �������������ҳ��Default.asp�������·��</td>
      <td colspan="3" class="tdbg"> <input name="SaveUpFilesPath" type="text" id="SaveUpFilesPath" value="<%=SaveUpFilesPath%>" size="30" maxlength="100"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>�������ϴ��ļ����ͣ�</strong><br>
        ֻ������չ����ÿ���ļ������á�|���ŷֿ���</td>
      <td colspan="3" class="tdbg"> <input name="UpFileType" type="text" id="UpFileType2" value="<%=UpFileType%>" size="50" maxlength="255"></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>ɾ������ʱ�Ƿ�ͬʱɾ�������е��ϴ��ļ���</strong><br>
        �˹�����ҪFSO֧�֡�</td>
      <td colspan="3" class="tdbg"> <input type="radio" name="DelUpFiles" value="Yes" <%if DelUpFiles="Yes" then response.write "checked"%>>
        �� &nbsp;&nbsp;&nbsp;&nbsp; <input type="radio" name="DelUpFiles" value="No" <%if DelUpFiles="No" then response.write "checked"%>>
        ��</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg"><strong>Session�Ự�ı���ʱ�䣺</strong><br>
        ��Ҫ���ں�̨����Ա��¼��Ϊ�˰�ȫ���벻Ҫ��ʱ�����̫����������Ϊ10����</td>
      <td colspan="3" class="tdbg"> <input name="SessionTimeout" type="text" id="SessionTimeout" value="<%=SessionTimeout%>" size="6" maxlength="5">
        ����</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td class="tdbg">&nbsp;</td>
      <td colspan="3" class="tdbg">&nbsp;</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" class="topbg"><strong>�ʼ�������ѡ��</strong></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>�ʼ����������</strong><br>
        ��һ��Ҫѡ����������Ѱ�װ�����</td>
      <td colspan="3" class="tdbg"> <select name="MailObject" id="MailObject">
          <option value="Jmail" selected>Jmail</option>
        </select> </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP��������ַ��</strong><br>
        ���������ʼ���SMTP������<br>
        ����㲻����˲������壬����ϵ��Ŀռ��� </td>
      <td colspan="3" class="tdbg"> <input name="MailServer" type="text" id="MailServer" value="<%=MailServer%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP��¼�û�����</strong><br>
        ����ķ�������ҪSMTP������֤ʱ�������ô˲���</td>
      <td colspan="3" class="tdbg"> <input name="MailServerUserName" type="text" id="MailServerUserName" value="<%=MailServerUserName%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP��¼���룺</strong><br>
        ����ķ�������ҪSMTP������֤ʱ�������ô˲��� </td>
      <td colspan="3" class="tdbg"> <input name="MailServerPassWord" type="password" id="MailServerPassWord" value="<%=MailServerPassWord%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="224" class="tdbg"><strong>SMTP����</strong>��<br>
        ����á�name@domain.com���������û�����¼ʱ����ָ��domain.com</td>
      <td colspan="3" class="tdbg"> <input name="MailDomain" type="text" id="MailDomain" value="<%=MailDomain%>" size="40"> 
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="4" align="center" class="tdbg"> <input name="Action" type="hidden" id="Action" value="SaveConfig"> 
        <input name="cmdSave" type="submit" id="cmdSave" value=" �������� " <% If ObjInstalled=false Then response.write "disabled" %>> 
      </td>
    </tr>
    <%
If ObjInstalled=false Then
	Response.Write "<tr class='tdbg'><td height='40' colspan='3'><b><font color=red>��ķ�������֧�� FSO(Scripting.FileSystemObject)! ����ʹ�ñ����ܡ�<br>��ֱ���޸ġ�Inc/config.asp���ļ��е����ݡ�</font></b></td></tr>"
End If
%>
  </table>
<%
end sub
%>
</form>
</body>
</html>
<%
sub SaveConfig()
	If ObjInstalled=false Then
		FoundErr=True
		ErrMsg=ErrMsg & "<br><li>��ķ�������֧�� FSO(Scripting.FileSystemObject)! </li>"
		exit sub
	end if
	dim fso,hf
	set fso=Server.CreateObject("Scripting.FileSystemObject")
	set hf=fso.CreateTextFile(Server.mappath("../inc/config.asp"),true)
	hf.write "<" & "%" & vbcrlf
	hf.write "Const SiteName=" & chr(34) & trim(request("SiteName")) & chr(34) & "        '��վ����" & vbcrlf
	hf.write "Const SiteTitle=" & chr(34) & trim(request("SiteTitle")) & chr(34) & "        '��վ����" & vbcrlf
	hf.write "Const SiteUrl=" & chr(34) & trim(request("SiteUrl")) & chr(34) & "        '��վ��ַ" & vbcrlf
	hf.write "Const EnterpriseMail=" & chr(34) & trim(request("EnterpriseMail")) & chr(34) & "        '��ҵ�ʾ�" & vbcrlf
	hf.write "Const LogoUrl=" & chr(34) & trim(request("LogoUrl")) & chr(34) & "        'Logo��ַ" & vbcrlf
	hf.write "Const BannerUrl=" & chr(34) & trim(request("BannerUrl")) & chr(34) & "        'Banner��ַ" & vbcrlf
	hf.write "Const High=" & trim(request("High")) & "        '�߶�" & vbcrlf
	hf.write "Const WebmasterName=" & chr(34) & trim(request("WebmasterName")) & chr(34) & "        'վ������" & vbcrlf
	hf.write "Const WebmasterEmail=" & chr(34) & trim(request("WebmasterEmail")) & chr(34) & "        'վ������" & vbcrlf
	hf.write "Const Copyright=" & chr(34) & trim(request("Copyright")) & chr(34) & "        '��Ȩ��Ϣ" & vbcrlf
	
	hf.write "Const MaxPerPage_Default=" & trim(request("MaxPerPage_Default")) & "        '��ҳÿҳ��Ʒ��������" & vbcrlf
	hf.write "Const New_count=" & trim(request("New_count")) &  "        '����������Ѷ����" & vbcrlf
	hf.write "Const Product_count=" & trim(request("Product_count")) &  "        '�����Ʒ�б���" & vbcrlf
	hf.write "Const MaxPerPage_Search=" & trim(request("MaxPerPage_Search")) & "        '��������ҳÿҳ������" & vbcrlf
	
	hf.write "Const MaxPerPage_Content=" & trim(request("MaxPerPage_Content")) &  "        'ÿҳ��ʾ��Լ�ַ���" & vbcrlf
	hf.write "Const NewsComment=" & chr(34) & trim(request("NewsComment")) & chr(34) & "        '�Ƿ������������۹���" & vbcrlf	
	hf.write "Const EnableArticleCheck=" & chr(34) & trim(request("EnableArticleCheck")) & chr(34) & "        '�Ƿ�����������˹���" & vbcrlf
	hf.write "Const EnableUploadFile=" & chr(34) & trim(request("EnableUploadFile")) & chr(34) & "        '�Ƿ񿪷��ļ��ϴ�" & vbcrlf
	hf.write "Const PopAnnounce=" & chr(34) & trim(request("PopAnnounce")) & chr(34) & "        '�Ƿ񵯳����洰��" & vbcrlf
	hf.write "Const HitsOfHot=" & trim(request("HitsOfHot")) & "        '�������µ����" & vbcrlf
	
	hf.write "Const MaxFileSize=" & trim(request("MaxFileSize")) & "        '�ϴ��ļ���С����" & vbcrlf
	hf.write "Const SaveUpFilesPath=" & chr(34) & trim(request("SaveUpFilesPath")) & chr(34) & "        '����ϴ��ļ���Ŀ¼" & vbcrlf
	hf.write "Const UpFileType=" & chr(34) & trim(request("UpFileType")) & chr(34) & "        '�������ϴ��ļ�����" & vbcrlf
	hf.write "Const DelUpFiles=" & chr(34) & trim(request("DelUpFiles")) & chr(34) & "        'ɾ������ʱ�Ƿ�ͬʱɾ�������е��ϴ��ļ�" & vbcrlf
	hf.write "Const SessionTimeout=" & trim(request("SessionTimeout")) & "        'Session�Ự�ı���ʱ��" & vbcrlf
	
	hf.write "Const MailObject=" & chr(34) & trim(request("MailObject")) & chr(34) & "        '�ʼ��������" & vbcrlf
	hf.write "Const MailServer=" & chr(34) & trim(request("MailServer")) & chr(34) & "        '���������ʼ���SMTP������" & vbcrlf
	hf.write "Const MailServerUserName=" & chr(34) & trim(request("MailServerUserName")) & chr(34) & "        '��¼�û���" & vbcrlf
	hf.write "Const MailServerPassWord=" & chr(34) & trim(request("MailServerPassWord")) & chr(34) & "        '��¼����" & vbcrlf
	hf.write "Const MailDomain=" & chr(34) & trim(request("MailDomain")) & chr(34) & "        '����" & vbcrlf
	hf.write "%" & ">"
	hf.close
	set hf=nothing
	set fso=nothing	
end sub
%>