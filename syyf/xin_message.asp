
<%
'==处理留言
if Request("action")=1 then
	if Trim(Request("name"))="" or Trim(Request("info"))="" then
		Response.write "<script language='javascript'>alert('信息填写不完整：\n姓名、留言内容必须填写！');history.go(-1);</script>"
		Response.End()
	end if
	
	sql="select * from xin_Message"
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open sql,conn,1,3
	
	rs.addnew
	rs("name")=nohtml(Trim(Request("name")))
	rs("tel")=nohtml(Trim(Request("tel")))
	rs("mail")=nohtml(Trim(Request("mail")))
	rs("info")=nohtml(Trim(Request("info")))
	rs("addtime")=now()
	rs("audit")=0
	rs.update
	
	rs.close
	set rs=nothing
	 Response.Write "<script language='javascript'>alert('留言成功,请等待管理员审核与回复!');document.location.href('message.asp');</script>"
	Response.End()
end if



'==读取留言列表

sql="select * from xin_Message where audit=1 order by id desc"
messagelist=""
set rs=server.CreateObject("ADODB.RECORDSET")
rs.open sql,conn,1,1
if not rs.eof then
pages = 5 '定义每页显示的记录数
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

if rs("re_info")="无回复" then
	messagelist=messagelist&"<div class=""gbtop""><ul><li>姓名："&rs("name")&"</li> &nbsp;<li>时间："&rs("addtime")&"</li></ul></div>" & "<div class=""gbinfo"">"&rs("info")&"</div>"&"<div class=""gbc""></div>" 
	
else
	messagelist=messagelist&"<div class=""gbtop""><ul><li>姓名："&rs("name")&"</li> &nbsp;<li>时间："&rs("addtime")&"</li></ul></div>" & "<div class=""gbinfo"">"&rs("info")&"</div>"&"<div class=""gbhf"">管理员回复："&rs("re_info")&"</div>" & "<div class=""gbc""></div>"
end if	
pages = pages - 1
rs.MoveNext
loop
end if
rs.Close
Set rs=Nothing

'== 构造分页
pager=	"<a href='?page=1&id="&request("id")&"'>首页</a> &nbsp;"&_
		"<a href='?page="&page-1&"&id="&request("id")&"'>上一页</a> &nbsp;"&_
		"<a href='?page="&page+1&"&id="&request("id")&"'>下一页</a> &nbsp;"&_
		"<a href='?page="&allpages&"&id="&request("id")&"'>尾页</a> &nbsp;"&_
		"页次："&page&"/"&allpages&"页"

%>