
<%
'==��ȡ�����б�
id=Request("id")
if id="" then
	sql="select * from xin_New order by id desc"
	menuname="��������"
else
	sql="select * from xin_New where lid="&id&" or lid in (Select id from xin_New_Pros where pid="&id&") order by id desc"
	sql2="select * from xin_New_Pros where id="&id
	set rs1=conn.execute(sql2)
	if not rs1.eof then
		menuname=rs1("name")
	else
		menuname="û�д˷���"
	end if
end if
newslist="<ul>"
set rs=server.CreateObject("ADODB.RECORDSET")
rs.open sql,conn,1,1
if not rs.eof then
pages = 13 '����ÿҳ��ʾ�ļ�¼��
rs.pageSize = pages '����ÿҳ��ʾ�ļ�¼��
allPages = rs.pageCount'����һ���ֶܷ���ҳ
page = Request.QueryString("page")'ͨ����������ݵ�ҳ��
'if������ڻ������Ŵ���
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

'== �����ҳ
pager=	"<a href='?page=1&id="&request("id")&"'>��ҳ</a> &nbsp;"&_
		"<a href='?page="&page-1&"&id="&request("id")&"'>��һҳ</a> &nbsp;"&_
		"<a href='?page="&page+1&"&id="&request("id")&"'>��һҳ</a> &nbsp;"&_
		"<a href='?page="&allpages&"&id="&request("id")&"'>βҳ</a> &nbsp;"&_
		"ҳ�Σ�"&page&"/"&allpages&"ҳ"

%>