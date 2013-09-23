<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="33" background="Images/l_mu_bg2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="117" height="33" background="Images/l_mu_bg1.gif" class="tt" style="background-repeat:no-repeat">留言管理</td>
        <td width="646" height="33">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td>
	<%
	id=Request("id")
	if id="" then
	lname="所有留言"
	sql="select * from xin_Message order by id desc"
	elseif id=1 then
	lname="未回复留言"
	sql="select * from xin_Message where re_info='无回复' order by id desc"
	elseif id=2 then
	lname="未审核留言"
	sql="select * from xin_Message where audit<>1 order by id desc"
	end if
	set rs=server.CreateObject("ADODB.RECORDSET")
	rs.open sql,conn,1,1
	if not rs.eof then
	pages = 10 '定义每页显示的记录数
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
	%>
	<table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
      <tr>
        <td width="87%" height="25" align="left" valign="middle" class="t2"><%=lname%></td>
        </tr>
      <tr>
        <td height="25" align="right" valign="middle" class="tb"><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
          <tr>
            <td width="12%" height="25" align="center" valign="middle" class="tb">姓名</td>
            <td width="15%" align="center" valign="middle" class="tb">电话</td>
            <td width="20%" align="center" valign="middle" class="tb">邮箱</td>
            <td width="15%" align="center" valign="middle" class="tb">时间</td>
            <td width="18%" align="center" valign="middle" class="tb">状态</td>
            <td width="20%" height="25" align="center" valign="middle" class="tb">管理</td>
          </tr>
          <tr>
            <td height="25" colspan="6" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
             <%
			 do while not rs.eof and pages > 0
			 %>
	          <tr>
                <td width="12%" height="30" align="center" valign="middle"><%=rs("name")%></td>
                <td width="15%" height="30" align="center" valign="middle"><%=rs("tel")%></td>
                <td width="20%" height="30" align="center" valign="middle"><%=rs("mail")%></td>
                <td width="15%" height="30" align="center" valign="middle"><%=rs("addtime")%></td>
                <td width="18%" height="30" align="center" valign="middle">
				<%
				if rs("audit")=0 then
				audit="<font color='red'>未审核</font>"
				else
				audit="已审核"
				end if
				if rs("re_info")="无回复" then
				reinfo="<font color='red'>未回复</font>"
				else
				reinfo="已回复"
				end if
				Response.Write(audit&","&reinfo)
				%>				</td>
                <td width="20%" height="30" align="center" valign="middle"><a href="Message_Edit.asp?id=<%=rs("id")%>">审核、回复</a> | <a href="Message_Del.asp?id=<%=rs("id")%>" ONCLICK="javascript:return confirm('提示：您确定要继续执行该操作吗？')">删除</a></td>
              </tr>
			  <%
			  pages = pages - 1
			  rs.movenext
			  loop
			  else
			  Response.Write("<font color='Red'>暂无内容！</font>")
			  End if
			  %>
            </table></td>
            </tr>
        </table></td>
        </tr>
      <tr>
        <td height="40" align="right" valign="middle" class="tb"><a href="?page=1&id=<%=Request("id")%>">首页</a>&nbsp; <a href="?page=<%=page-1%>&id=<%=Request("id")%>">上一页</a> &nbsp; <a href="?page=<%=page+1%>&id=<%=Request("id")%>">下一页</a>&nbsp; <a href="?page=<%=allpages%>&id=<%=Request("id")%>">尾页</a> &nbsp;
      页次：<%=page%>/<%=allpages%>页</td>
      </tr>
    </table>
	<%
	rs.close
	set rs=nothing
	%>
	</td>
  </tr>
</table>
</body>
</html>
