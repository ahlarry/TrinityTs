 <%
'����Ԫ��
Dim fsoObject 			'�ļ�ϵͳ����
Dim tsObject 			'����ϵͳ����
Dim filObject			'�ļ�����
Dim lngVisitorNumber 		'��������������
Dim intDisplayDigitsLoopCount 	'ѭ��������ʾ


On Error Resume Next

lngVisitorNumber = 0
	
Set fsoObject = Server.CreateObject("Scripting.FileSystemObject")

Set filObject = fsoObject.GetFile(Server.MapPath("visitor_counter.txt"))
	
Set tsObject = filObject.OpenAsTextStream
	
lngVisitorNumber = CLng(tsObject.ReadAll)

	
If Session("lngSessionVisitorNum") = "" Then
		
	'Increment the visitor counter number by 1
	lngVisitorNumber = lngVisitorNumber + 1
		
	'Place the Visitor number in the session visitor number
	Session("lngSessionVisitorNum") = lngVisitorNumber
	
Else
	
	'Place the Visitor number in the session visitor number
	Session("lngSessionVisitorNum") = lngVisitorNumber
End if
	
			
Set tsObject = fsoObject.CreateTextFile((Server.MapPath("visitor_counter.txt")), True)
	
tsObject.Write CStr(lngVisitorNumber)

Set fsoObject = Nothing
Set tsObject = Nothing
Set filObject = Nothing

'************����������޸�����һ���еļ���������˵��************
Response.Write "		<font size=2></font>"


For intDisplayDigitsLoopCount = 1 to Len(lngVisitorNumber)
	'Display the graphical hit count by getting the path to the image using the mid function
'**********����������޸�����һ��src=""counter_images/11/�еļ�����ͼƬ�ļ���·�����޸�Ϊsrc=""counter_images/3/��������ϲ��������ͼƬ�����ļ���************
	Response.Write "<img src=""counter_images/6/" & Mid(lngVisitorNumber, intDisplayDigitsLoopCount, 1) & ".gif"">"
Next

%>
