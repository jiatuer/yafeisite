<!-- #include file="Include/Chk.asp" -->
<!-- #include file="../Include/conn.asp" -->
<%
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('��������!');document.location.href('Link_Manage.asp');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from xin_link where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('���ݿ���û�иü�¼��');document.location.href('Link_Manage.asp');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('�ɹ�ɾ��ָ������Ϣ��');document.location.href('Link_Manage.asp');</script>"
end if
%>