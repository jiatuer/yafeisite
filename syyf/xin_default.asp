<%
'== 读取最新新闻
sql="select top 8 Id,titles,lid,Addtime from xin_New  Where lid=45 order by id desc"
newslist="<ul>"
Set rs=conn.execute(sql)
do while not rs.eof
if len(rs("titles"))>19 then
	titles=left(rs("titles"),19)&"..."
else
	titles=rs("titles")
end if
	newslist=newslist&"<li><a href='news_view.asp?id="&rs("id")&"' target='_blank'>"&titles&"</a></li>"
rs.MoveNext
loop
newslist=newslist&"</ul>"
rs.Close
Set rs=Nothing


'== 读取室外设计

sql="select top 6 * from xin_show order by id desc"
showlist="<ul>"
Set rs=conn.execute(sql)
do while not rs.eof

	showlist=showlist&"<li><a href='show_view.asp?id="&rs("id")&"' target='_blank'><img src='./uploadfiles/upLoadImages/"&rs("img")&"' /></a><span><a href='show_view.asp?id="&rs("id")&"'>"&rs("title")&"</a></span></li>"

rs.MoveNext
loop
showlist=showlist&"</ul>"
rs.Close
Set rs=Nothing

'== 读取室内设计
sql="select top 6 * from xin_product order by id desc"
productlist="<ul>"
Set rs=conn.execute(sql)
do while not rs.eof

	productlist=productlist&"<li><a href='product_view.asp?id="&rs("id")&"' target='_blank'><img src='./uploadfiles/upLoadImages/"&rs("img")&"' /></a><span><a href='product_view.asp?id="&rs("id")&"'>"&rs("title")&"</a></span></li>"

rs.MoveNext
loop
productlist=productlist&"</ul>"
rs.Close
Set rs=Nothing
%>