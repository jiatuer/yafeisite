<!-- #include file="../Include/Chk.asp" -->
<head>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	font-size:12px;
	background-color: #ffffff;
}
-->
</style>


<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../Images/main.css" rel="stylesheet" type="text/css">
</head>
<body>
<!--#include FILE="upload.inc"-->
<%
dim upload,file,formName,formPath,iCount,filename,fileExt
set upload=new upload_5xSoft ''�����ϴ�����
formPath="../../uploadfiles/uploadimages/"
 ''��Ŀ¼���(/)
if right(formPath,1)<>"/" then formPath=formPath&"/" 
iCount=0
for each formName in upload.file ''�г������ϴ��˵��ļ�
 set file=upload.file(formName)  ''����һ���ļ�����
 if file.filesize<100 then
 	response.write "����ѡ����Ҫ�ϴ���ͼƬ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]"
	response.end
 end if
 	
 if file.filesize>1000000 then
 	response.write "<font size=3><br>ͼƬ��С���������ơ�[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if

 fileExt=lcase(right(file.filename,4))

 if fileEXT<>".jpg" and fileEXT<>".gif" then
 	response.write "<font size=3><br>�ļ���ʽֻ��Ϊjpg��gif��ʽ��[ <a href=# onclick=history.go(-1)>�����ϴ�</a> ]</font>"
	response.end
 end if 
 randomize
 ranNum=int(90000*rnd)+10000
 filename=formPath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 filename1=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&fileExt
 
 if file.FileSize>0 then         ''��� FileSize > 0 ˵�����ļ�����
 file.SaveAs Server.mappath(filename)   ''�����ļ�
response.write "<script>parent.form1.img.value='"&FileName1&"'</script>"

  iCount=iCount+1
 end if
 set file=nothing
next
set upload=nothing  ''ɾ���˶���

Response.Write "<script>parent.form1.img.value='"&FileName1&"';</script>"

Response.Write("�ϴ��ɹ���")
%>
</body>
</html>