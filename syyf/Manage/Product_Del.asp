<!-- #include file="Include/Chk.asp" -->
<!-- #include file="../Include/conn.asp" -->
<%
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('��������!');document.location.href('Product_Manage.asp');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from xin_Product where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('���ݿ���û�иü�¼��');document.location.href('Product_Manage.asp');</script>"
	Response.End()
else
	if rs("img")<>"noimages.gif" then
		DeleFile("upload/uploadimages/"&rs("img"))
	end if
	
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('�ɹ�ɾ��ָ������Ϣ��');document.location.href('Product_Manage.asp');</script>"
end if

'====
Function DeleFile(FilePath)
    On Error Resume Next
    Set Del=Server.CreateObject("Scripting.FileSystemObject")
    if Err <> 0 Then 
        DelFile="�ÿռ䲻֧��FSO������޷�ɾ���ļ���"
    else
        if InStr(FilePath, ",") > 0 then
            FilePath=Split(FilePath,",")
            For i = 0 to ubound(FilePath)
                If Del.FileExists(Server.MapPath(FilePath))=True Then Del.DeleteFile Server.MapPath(FilePath(i)),true
            Next
        Else
            If Del.FileExists(Server.MapPath(FilePath))=True Then Del.DeleteFile Server.MapPath(FilePath),true
        End if
    End if
    Set Del=Nothing
End Function
%>