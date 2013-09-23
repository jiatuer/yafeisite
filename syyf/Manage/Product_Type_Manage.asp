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
        <td width="117" height="33" background="Images/l_mu_bg1.gif" class="tt" style="background-repeat:no-repeat">产品管理</td>
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
	sql="select * from xin_Product_Info where pid=0 order by ord asc"
	set rs=server.CreateObject("ADODB.RECORDSET")
	rs.open sql,conn,1,1
	if not rs.eof then
	pages = 8 '定义每页显示的记录数
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
        <td width="87%" height="25" align="left" valign="middle" class="t2">栏目列表</td>
        </tr>
      <tr>
        <td height="25" align="right" valign="middle" class="tb"><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
          <tr>
            <td width="30%" height="25" align="center" valign="middle" class="tb">分类名称</td>
            <td width="30%" align="center" valign="middle" class="tb">显示顺序(数值越小越靠前)</td>
            <td width="40%" height="25" align="center" valign="middle" class="tb">管理</td>
          </tr>
          <tr>
            <td height="25" colspan="3" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
             <%
			 do while not rs.eof and pages > 0
			 %>
	          <tr>
                <td width="30%" height="25" align="center" valign="middle"><%=rs("title")%></td>
                <td width="30%" align="center" valign="middle"><%=rs("ord")%></td>
                <td width="40%" height="25" align="center" valign="middle"><a href="Product_Type_Edit.asp">修改</a> | <a href="Product_Type_Del.asp?id=<%=rs("id")%>" ONCLICK="javascript:return confirm('提示：您确定要继续执行该操作吗？')">删除</a>&nbsp;[<a href="product_type_add.asp?pid=<%=rs("id")%>">添加二级分类</a>]</td>
              </tr>
			  <%
			  sql1="select * from xin_Product_Info where pid="&rs("id")&" order by ord asc"
			  set js=server.CreateObject("ADODB.Recordset")
			  js.open sql1,conn,1,1
			  do while not js.eof
			  %>
	          <tr>
	            <td height="30" align="center" valign="middle" class="ej">--&gt;&nbsp;<%=js("name")%></td>
	            <td height="30" align="center" valign="middle" class="ej">--&gt;&nbsp;<%=js("ord")%></td>
	            <td height="30" align="center" valign="middle" class="ej">--&gt;&nbsp;[<a href="Product_Type_Edit.asp">修改</a>]&nbsp;[<a href="Product_Type_Del.asp?id=<%=js("id")%>" ONCLICK="javascript:return confirm('提示：您确定要继续执行该操作吗？')">删除</a>]</td>
	            </tr>
			  <%
			  js.movenext
			  loop
			  js.close
			  set js=nothing
			  %>	
			  <%
			  pages = pages - 1
			  rs.movenext
			  loop
			  else
			  Response.Write("<font color='Red'>数据库暂无内容！</font>")
			End if
			  %>
            </table></td>
            </tr>
        </table></td>
        </tr>
      <tr>
        <td height="40" align="right" valign="middle" class="tb"><a href="?page=1">首页</a>&nbsp; <a href="?page=<%=page-1%>">上一页</a> &nbsp; <a href="?page=<%=page+1%>">下一页</a>&nbsp; <a href="?page=<%=allpages%>">尾页</a> &nbsp;
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
