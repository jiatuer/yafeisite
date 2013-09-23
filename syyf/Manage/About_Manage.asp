<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
<style type="text/css">
<!--
.STYLE1 {color: #FF0000}
-->
</style>
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="33" background="Images/l_mu_bg2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="117" height="33" background="Images/l_mu_bg1.gif" class="tt" style="background-repeat:no-repeat">单页管理</td>
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
	sql="select * from xin_About order by tid asc,ord asc"
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
        <td width="87%" height="25" align="left" valign="middle" class="t2">单页列表 <span class="STYLE1">说明：点右侧的修改，可以修改相应的单页内容！</span></td>
        </tr>
      <tr>
        <td height="25" align="right" valign="middle" class="tb"><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
          <tr>
            <td width="25%" height="25" align="center" valign="middle" class="tb">名称</td>
            <td width="25%" align="center" valign="middle" class="tb">显示位置(0:头部导航 1:底部导航)</td>
            <td width="25%" align="center" valign="middle" class="tb">顺序(数值越小越靠前)</td>
            <td width="25%" height="25" align="center" valign="middle" class="tb">管理</td>
          </tr>
          <tr>
            <td height="25" colspan="4" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
             <%
			 do while not rs.eof and pages > 0
			 %>
	          <tr>
                <td width="25%" height="25" align="center" valign="middle"><%=rs("name")%></td>
                <td width="25%" align="center" valign="middle"><%=rs("tid")%></td>
                <td width="25%" align="center" valign="middle"><%=rs("ord")%></td>
                <td width="25%" height="25" align="center" valign="middle"><a href="About_Edit.asp?id=<%=rs("id")%>">修改</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; |&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="About_Del.asp?id=<%=rs("id")%>" ONCLICK="javascript:return confirm('提示：您确定要继续执行该操作吗？')">删除</a></td>
              </tr>
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
