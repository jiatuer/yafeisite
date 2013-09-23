<%
'==读取产品内容
id=Request("id")
if id="" or not isnumeric(id) then
Response.Write("<script>alert('访问参数错误！');document.location.href('index.asp');</script>")
Response.End()
end if

sql="select * from xin_Product where id="&id
set rs=server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3

if rs.eof then
Response.Write("<script>alert('参数错误！');document.location.href('index.asp');</script>")
Response.End()
end if
rs("hit")=rs("hit")+1
rs.update
productname=rs("title")
producttime=datevalue(rs("addtime"))
producthit=rs("hit")
productinfo=rs("info")
productfrom=rs("from")
productauthor=rs("author")
productimg="<img src='./UpLoadfiles/upLoadImages/"&rs("img")&"' />"
rs.close
set rs=nothing
%>