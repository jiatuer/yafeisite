<%
'==��ȡ��Ʒ����
id=Request("id")
if id="" or not isnumeric(id) then
Response.Write("<script>alert('���ʲ�������');document.location.href('index.asp');</script>")
Response.End()
end if

sql="select * from xin_Show where id="&id
set rs=server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3

if rs.eof then
Response.Write("<script>alert('��������');document.location.href('index.asp');</script>")
Response.End()
end if
rs("hit")=rs("hit")+1
rs.update
showname=rs("title")
showtime=datevalue(rs("addtime"))
showhit=rs("hit")
showinfo=rs("info")
showauthor=rs("author")
showfrom=rs("from")
showimg="<img src='./UpLoadfiles/upLoadImages/"&rs("img")&"' />"
rs.close
set rs=nothing
%>