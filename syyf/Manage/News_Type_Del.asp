<!-- #include file="Include/Chk.asp" -->
<!-- #include file="../Include/conn.asp" -->
<%
	id=request("id")
	if id="" then
	Response.Write "<script language='javascript'>alert('参数错误!');document.location.href('News_Type_Manage.asp');</script>"
	Response.End()
	end if
set rs=server.createobject("adodb.recordset")
rs.open "Select * from xin_New_Pros where id="&Request("id"),conn,1,3
if rs.bof and rs.eof then
	Response.Write "<script language='javascript'>alert('数据库中没有该记录！');document.location.href('News_Type_Manage.asp');</script>"
	Response.End()
else
	rs.Delete
	rs.Update
	Response.Write "<script language='javascript'>alert('成功删除指定的信息！');document.location.href('News_Type_Manage.asp');</script>"
end if
%>