<%
'==��ȡ��������
id=Request("id")
if id="" or not isnumeric(id) then
Response.Write("<script>alert('���ʲ�������');document.location.href('index.asp');</script>")
Response.End()
end if

sql="select * from xin_New where id="&id
set rs=server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3

if rs.eof then
Response.Write("<script>alert('��������');document.location.href('index.asp');</script>")
Response.End()
end if
rs("hit")=rs("hit")+1
rs.update
newstitle=rs("titles")
newstime=datevalue(rs("addtime"))
newshit=rs("hit")
newsfrom=rs("from")
newsinfo=rs("info")
newsauthor=rs("author")
rs.close
set rs=nothing
%>