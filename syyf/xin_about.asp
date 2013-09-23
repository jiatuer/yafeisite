<%
'== 读取单页信息
id=Request("id")
if id="" or not isnumeric(id) then
Response.Write("<script>alert('访问参数错误！');document.location.href('index.asp');</script>")
Response.End()
end if

sql="select * from xin_About where id="&id
Set rs=conn.execute(sql)

	pagename=rs("name")
	pageinfo=rs("info")
	url=rs("url")

rs.Close
Set rs=Nothing

if url<>"" then Response.Redirect(url)


'== 重新读取主导航
sql1="select * from xin_About where tid=0 order by ord asc"
mainnav=""
Set nrs=conn.execute(sql1)
do while not nrs.eof
	if cstr(nrs("id"))=cstr(id) then
		mainnav=mainnav&"<li><a href='about.asp?id="&nrs("id")&"'><img src='Images/icon5x.png' width='11' height='11' border='0' />&nbsp;&nbsp;<b>"&nrs("name")&"</b></a></li>"
	else
		mainnav=mainnav&"<li><a href='about.asp?id="&nrs("id")&"'><img src='Images/icon5x.png' width='11' height='11' border='0' />&nbsp;&nbsp;<b>"&nrs("name")&"</b></a></li>"
	end if
nrs.MoveNext
loop
nrs.Close
Set nrs=Nothing

%>