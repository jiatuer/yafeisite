<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!-- #include file="Include/MD5.asp" -->
<%
if Request("Action")=1 then
	if Trim(Request("name"))="" or Trim(Request("pass"))="" then
				Response.write "<script language='javascript'>alert('用户名或密码不能为空！');history.go(-1);</script>"
		Response.End()
	end if
	SQL="Select * from xin_Admin where name='"&Trim(Request("name"))&"'"
	set rs=server.createobject("adodb.recordset")
	rs.open SQL,conn,1,3
	if not rs.eof then
		Response.Write "<script language='javascript'>alert('已存在同名管理员,请重新输入!');history.back();</script>"
		Response.End()
	end if
	rs.AddNew
	rs("name")=Trim(Request("name"))
	rs("pass")=MD5(Trim(Request("pass")))

	rs.Update
	rs.Close
	Set rs=nothing
	Response.Write "<script language='javascript'>alert('您成功添加了一位管理员!');document.location.href('Admin_Manage.asp');</script>"
	Response.End()
else
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>
<link href="Images/main.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="ta">
      <tr>
        <td width="50%" height="25" align="left" valign="middle" class="t2">添加管理员</td>
        <td width="50%" align="right" valign="middle" class="t2">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="8"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
      <tr>
        <td colspan="2" align="left" valign="top" bgcolor="#FFFFFF">
          <table width="50%" border="0" cellspacing="0" cellpadding="0">
		  <form name="form1" method="post" action="?action=1">
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">用户名：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="name" type="text" class="main" id="name" size="20" maxlength="20"></td>
            </tr>
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">密&nbsp;&nbsp;码：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="pass" type="text" class="main" id="pass" size="20" maxlength="20"></td>
            </tr>
            <tr>
              <td height="50" colspan="2" align="center" valign="middle"><input name="Submit" type="submit" class="inputbut" value="提交">
                &nbsp; <input name="Submit2" type="reset" class="inputbut" value="重置"></td>
              </tr>
			</form>
          </table>
                
        </td>
        </tr>
      
    </table></td>
  </tr>
</table>
</body>
</html>
<%
end if
%>