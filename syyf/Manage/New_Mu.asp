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
        <td width="117" height="33" background="Images/l_mu_bg1.gif" class="tt" style="background-repeat:no-repeat">新闻管理</td>
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
	if Request("lid")="" or Request("lid")="error" then
	Response.Write "<script language='javascript'>alert('请选择栏目！');history.go(-1);</script>"
	Response.End()
	end if

	sql="select * from xin_Index where lid="&Request("lid")&" order by id desc"
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
        <td width="87%" height="25" colspan="2" align="left" valign="middle" class="t2">新闻列表</td>
        </tr>
      <tr>
        <td height="25" colspan="2" align="right" valign="middle" class="tb"><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
          <tr>
            <td width="60%" height="25" align="center" valign="middle" class="tb">新闻标题</td>
            <td width="20%" align="center" valign="middle" class="tb">所属栏目</td>
            <td width="20%" height="25" align="center" valign="middle" class="tb">管理</td>
          </tr>
          <tr>
            <td height="25" colspan="3" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
             <%
			 do while not rs.eof and pages > 0
			 %>
	          <tr>
                <td width="60%" height="30" align="left" valign="middle">&nbsp;&nbsp;<a href="New_View.asp?id=<%=rs("id")%>" target="_blank"><%=rs("titles")%></a></td>
                <td width="20%" height="30" align="center" valign="middle">
				<%
				lid=rs("lid")
				sql1="select * from xin_Index_Type where id="&lid
				set rs1=server.CreateObject("adodb.recordset")
				rs1.open sql1,conn,1,1
				if rs1.eof then
				response.Write("<font color='Red'>栏目不存在</font>")
				else
				response.Write(rs1("name"))
				rs1.close
				set rs1=nothing
				end if
				%>				</td>
                <td width="20%" height="30" align="center" valign="middle"><a href="New_Edit.asp?id=<%=rs("id")%>">修改</a>| <a href="New_Del.asp?id=<%=rs("id")%>" ONCLICK="javascript:return confirm('提示：您确定要继续执行该操作吗？')">删除</a></td>
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
	  <form name="form1" method="post" action="New_Mu.asp">
        <td height="30" align="left" valign="middle" class="tb">
          &nbsp;<select name="lid" id="lid">
            <option value="error">--选择栏目查看--</option>
			<%
			sql2="select * from xin_Index_Type where pid=0"
			set rs2=server.CreateObject("adodb.recordset")
			rs2.open sql2,conn,1,1
			do while not rs2.eof
			response.Write"<option value='"& rs2("id") &"'>"& rs2("name") &"</option>"
			
				Set Rs3=Conn.Execute("select * from xin_Index_Type where pid="&Rs2("id")&"")
				do while not Rs3.eof
				Response.Write "<option style='color:#999999' value='" & Rs3("id") & "'>--"& Rs3("name") &"</option>"
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
        <td height="30" align="right" valign="middle" class="tb"><a href="?page=1&lid=<%=request("lid")%>">首页</a>&nbsp; <a href="?page=<%=page-1%>&lid=<%=request("lid")%>">上一页</a> &nbsp; <a href="?page=<%=page+1%>&lid=<%=request("lid")%>">下一页</a>&nbsp; <a href="?page=<%=allpages%>&lid=<%=request("lid")%>">尾页</a> &nbsp;
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
