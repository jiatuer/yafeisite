<%
'==读取新闻内容
id=Request("id")
if id="" or not isnumeric(id) then
Response.Write("<script>alert('访问参数错误！');document.location.href('index.asp');</script>")
Response.End()
end if

sql="select * from xin_Index where id="&id
set rs=server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3

if rs.eof then
Response.Write("<script>alert('参数错误！');document.location.href('index.asp');</script>")
Response.End()
end if
rs("hit")=rs("hit")+1
rs.update
newstitle=rs("titles")
newstime=rs("addtime")
newshit=rs("hit")
newsfrom=rs("from")
newsinfo=rs("info")
rs.close
set rs=nothing
%>