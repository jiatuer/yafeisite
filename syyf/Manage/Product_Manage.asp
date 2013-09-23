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
	sql="select * from xin_Product order by id desc"
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
        <td width="87%" height="25" colspan="2" align="left" valign="middle" class="t2">产品列表</td>
        </tr>
      <tr>
        <td height="25" colspan="2" align="right" valign="middle" class="tb"><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
          <tr>
            <td width="30%" height="25" align="center" valign="middle" class="tb">产品名称</td>
            <td width="30%" align="center" valign="middle" class="tb">所属分类</td>
            <td width="30%" height="25" align="center" valign="middle" class="tb">管理</td>
          </tr>
          <tr>
            <td height="25" colspan="3" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
             <%
			 do while not rs.eof and pages > 0
			 %>
	          <tr>
                <td width="30%" height="25" align="center" valign="middle"><%=rs("title")%></td>
                <td width="30%" height="25" align="center" valign="middle">
				<%
				tid=rs("tid")
				sql1="select * from xin_Product_Info where id="&tid
				set rs1=server.CreateObject("adodb.recordset")
				rs1.open sql1,conn,1,1
				if rs1.eof then
				response.Write("<font color='Red'>分类不存在</font>")
				else
				response.Write(rs1("title"))
				rs1.close
				set rs1=nothing
				end if
				%></td>
                <td width="30%" height="25" align="center" valign="middle"><a href="Product_Edit.asp?id=<%=rs("id")%>">修改</a> | <a href="Product_Del.asp?id=<%=rs("id")%>" ONCLICK="javascript:return confirm('提示：您确定要继续执行该操作吗？')">删除</a></td>
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
	  <form name="form1" method="post" action="Product_Mu.asp">
        <td height="30" align="left" valign="middle" class="tb">
          &nbsp;<select name="tid" id="tid">
            <option value="error">--选择栏目查看--</option>
			<%
			sql2="select * from xin_Product_Info"
			set rs2=server.CreateObject("adodb.recordset")
			rs2.open sql2,conn,1,1
			do while not rs2.eof
			response.Write"<option value='"&rs2("id")&"'>"&rs2("title")&"</option>"
			
				Set Rs3=Conn.Execute("select * from xin_Product_Info where pid="&Rs2("id")&"")
				do while not Rs3.eof
				Response.Write "<option style='color:#999999' value='" & Rs3("id") & "'>--"& Rs3("title") &"</option>"
				Rs3.movenext
				loop
				rs3.close
			
			rs2.movenext
			loop
			rs2.close
			set rs2=nothing
			%>
          </select>
          <input name="Submit2" type="submit" class="inputbut" value="列表">        </td>
		</form>
        <td height="30" align="right" valign="middle" class="tb"><a href="?page=1">首页</a>&nbsp; <a href="?page=<%=page-1%>">上一页</a> &nbsp; <a href="?page=<%=page+1%>">下一页</a>&nbsp; <a href="?page=<%=allpages%>">尾页</a> &nbsp;
      页次：<%=page%>/<%=allpages%>页</td>
      </tr>
    </table>
	<%
	rs.close
	set rs=nothing
	%>	</td>
  </tr>
</table>
</body>
</html>
