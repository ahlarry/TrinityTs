<%@language=vbscript codepage=936 %>
<!--#include file="Inc/config.asp"-->
<!--#include file="Inc/upfile_class.asp"-->
<%
const upload_type=0   '�ϴ�������0=�޾�������ϴ��࣬1=FSO�ϴ� 2=lyfupload��3=aspupload��4=chinaaspupload
dim upload,oFile,formName,SavePath,filename,fileExt
dim ImgWidth,ImgHeight,AlignType
dim EnableUpload
dim arrUpFileType
dim ranNum
dim msg,FoundErr
msg=""
FoundErr=false
EnableUpload=false
SavePath = SaveUpFilesPath    '����ϴ��ļ���Ŀ¼
if right(SavePath,1)<>"/" then SavePath=SavePath&"/" '��Ŀ¼���(/)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
</head>
<body leftmargin="2" topmargin="0" marginwidth="0" marginheight="0">
<%
if EnableUploadFile="No" then
	response.write "ϵͳδ�����ļ��ϴ�����"
else
	if session("AdminName")="" then
		response.Write("���¼����ʹ�ñ����ܣ�")
	else
		select case upload_type
			case 0
				call upload_0()  'ʹ�û���������ϴ���
			case else
				'response.write "��ϵͳδ���Ų������"
				'response.end
		end select
	end if
end if
%>
</body>
</html>
<%
sub upload_0()    'ʹ�û���������ϴ���
	set upload=new upfile_class ''�����ϴ�����
	upload.GetData(104857600)   'ȡ���ϴ�����,��������ϴ�100M
	if upload.err > 0 then  '�������
		select case upload.err
			case 1
				response.write "����ѡ����Ҫ�ϴ����ļ���"
			case 2
				response.write "���ϴ����ļ��ܴ�С������������ƣ�100M��"
		end select
		response.end
	end if
		
	ImgWidth=trim(upload.form("ImgWidth"))
	ImgHeight=trim(upload.form("ImgHeight"))
	AlignType=trim(upload.form("AlignType"))
	if ImgWidth="" then
		ImgWidth=0
	else
		ImgWidth=Clng(ImgWidth)
	end if
	if ImgHeight="" then
		ImgHeight=0
	else
		ImgHeight=Clng(ImgHeight)
	end if
	if AlignType="" then
		AlignType=0
	else
		AlignType=Clng(AlignType)
	end if
	
	for each formName in upload.file '�г������ϴ��˵��ļ�
		EnableUpload=False
		set ofile=upload.file(formName)  '����һ���ļ�����
	    if ofile.filesize<100 then
			msg="����ѡ����Ҫ�ϴ����ļ���"
			FoundErr=True
		end if
		if ofile.filesize>(MaxFileSize*1024) then
            msg="�ļ���С���������ƣ����ֻ���ϴ�" & CStr(MaxFileSize) & "K���ļ���"
			FoundErr=true			
		end if

		fileExt=lcase(ofile.FileExt)
		arrUpFileType=split(UpFileType,"|")
		for i=0 to ubound(arrUpFileType)
			if fileEXT=trim(arrUpFileType(i)) then
				EnableUpload=true
				exit for
			end if
		next
		if fileEXT="asp" or fileEXT="asa" or fileEXT="aspx" or fileEXT="cer" or fileEXT="cdx" then
			EnableUpload=false
		end if
		if EnableUpload=false then
			msg="�����ļ����Ͳ������ϴ���\n\nֻ�����ϴ��⼸���ļ����ͣ�" & nt2003.site_setting(24)
			FoundErr=true
		end if
		
		
		strJS="<SCRIPT language=javascript>" & vbcrlf
		if FoundErr<>true then
			randomize
			ranNum=int(900*rnd)+100
			filename=SavePath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&fileExt

			ofile.SaveToFile Server.mappath(FileName)   '�����ļ�

			msg="�ϴ��ļ��ɹ���"
			
			strJS=strJS & "parent.HtmlEdit.focus();" & vbcrlf 	  		
			strJS=strJS & "var range = parent.HtmlEdit.document.selection.createRange();" & vbcrlf
			select case FileExt
			    case "jpg","gif","png","bmp"
					strJS=strJS & "range.pasteHTML('<img src=" & filename
					if ImgWidth>0 then
						strJS=strJS & " width=" & ImgWidth
					end if
					if ImgHeight>0 then
						strJS=strJS & " height=" & ImgHeight
					end if
					select case AlignType
						case 0,1
							strJS=strJS & " align=left"
						case 2
							strJS=strJS & " align=center"
						case 3
							strJS=strJS & " align=right"
					end select
					strJS=strJS & " border=0>');" & vbcrlf
				case "swf"
					strJS=strJS & "range.pasteHTML('<object classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"
					if ImgWidth>0 then
						strJS=strJS & " width=" & ImgWidth
					end if
					if ImgHeight>0 then
						strJS=strJS & " height=" & ImgHeight
					end if
		  			strJS=strJS & "><param name=movie value=" & FileName & ">"
		  			strJS=strJS & "<param name=quality value=high>"
		  			strJS=strJS & "<embed src=" & FileName & " quality=high pluginspage=http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash type=application/x-shockwave-flash"
					if ImgWidth>0 then
						strJS=strJS & " width=" & ImgWidth
					end if
					if ImgHeight>0 then
						strJS=strJS & " height=" & ImgHeight
					end if
					strJS=strJS & "></embed></object>');" & vbcrlf
				case else
		  			strJS=strJS & "range.text='[upload=" & FileExt & "]" & FileName & "[/upload]';" & vbcrlf
	  		end select
			strJS=strJS & "parent.parent.AddItem('" & filename & "')" & vbcrlf
		end if
		strJS=strJS & "alert('" & msg & "');" & vbcrlf
	  	strJS=strJS & "history.go(-1);" & vbcrlf
		strJS=strJS & "parent.HtmlEdit.focus();" & vbcrlf 	  		
	  	strJS=strJS & "</script>"
	  	response.write strJS
		set file=nothing
	next
	set upload=nothing
end sub
%>