<!-- #include file="Include/Chk.asp" -->
<!-- #include file="../Include/Conn.asp" -->
<%
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('��������!');document.location.href('Admin_Manage.asp');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from xin_Admin where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('���ݿ���û�иü�¼��');document.location.href('Admin_Manage.asp');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('�ɹ�ɾ��ָ������Ϣ��');document.location.href('Admin_Manage.asp');</script>"
end if
%>