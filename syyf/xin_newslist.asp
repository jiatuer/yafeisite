
<%
'==读取新闻列表
id=Request("id")
if id="" then
	sql="select * from xin_New order by id desc"
	menuname="所有新闻"
else
	sql="select * from xin_New where lid="&id&" or lid in (Select id from xin_New_Pros where pid="&id&") order by id desc"
	sql2="select * from xin_New_Pros where id="&id
	set rs1=conn.execute(sql2)
	if not rs1.eof then
		menuname=rs1("name")
	else
		menuname="没有此分类"
	end if
end if
newslist="<ul>"
set rs=server.CreateObject("ADODB.RECORDSET")
rs.open sql,conn,1,1
if not rs.eof then
pages = 13 '定义每页显示的记录数
rs.pageSize = pages '定义每页显示的记录数
allPages = rs.pageCount'计算一共能分多少页
page = Request.QueryString("page")'通过浏览器传递的页数
'if语句属于基本的排错处理
if isEmpty(page) or Cint(page) < 1 then
page = 1
elseif Cint(page) > allPages then
page = allPages 
end if
rs.AbsolutePage = page
do while not rs.eof and pages > 0
if len(rs("titles"))>35 then
	titles=left(rs("titles"),35)&"..."
else
	titles=rs("titles")
end if

	newslist=newslist&"<li><a href='news_view.asp?id="&rs("id")&"'>&middot;&nbsp;"&titles&"&nbsp;&nbsp;&nbsp;["&datevalue(rs("addtime"))&"]</a></li>"
	
pages = pages - 1
rs.MoveNext
loop
end if
newslist=newslist&"</ul>"
rs.Close
Set rs=Nothing

'== 构造分页
pager=	"<a href='?page=1&id="&request("id")&"'>首页</a> &nbsp;"&_
		"<a href='?page="&page-1&"&id="&request("id")&"'>上一页</a> &nbsp;"&_
		"<a href='?page="&page+1&"&id="&request("id")&"'>下一页</a> &nbsp;"&_
		"<a href='?page="&allpages&"&id="&request("id")&"'>尾页</a> &nbsp;"&_
		"页次："&page&"/"&allpages&"页"

%>