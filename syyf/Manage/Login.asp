<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/MD5.asp" -->
<%
function R(s)
R = Replace(Trim(s), "'","")
R = Replace(R, """","")

end function
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>欢迎登录！</title>
<%
if Request("action")=1 then
name=R(Trim(Request("name")))
pass=MD5(R(Trim(Request("pass"))))

if name="" or pass="" then
	Response.Write("<script>alert('用户名或密码不能为空！');history.go(-1);</script>")
	Response.End()
end if

sql="select * from Xin_Admin where name='"&name&"' and pass='"&pass&"'"
set rs=server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	Response.Write("<script>alert('用户名或密码错误，请重新输入！');document.location.href('Login.asp');</script>")
	Response.End()
else
	Session("Admin")=rs("name")
	Response.Redirect("Manage.asp")
	
end if

end if
%>
<style type="text/css">
<!--
body {

	font-size:12px;
	color:#000000;
}
.cop{
font-size:12px;
color:#666666;
}
.t2 {
color:#FF6600;
font-size: 14px;
font-weight:bold;
}

.t1 {font-size: 20px;
     font-weight:bold;
}


.main{
	BORDER-RIGHT: #ccc 1px solid; PADDING-RIGHT: 0px; BORDER-TOP: #ccc 1px solid; PADDING-LEFT: 7px; BACKGROUND: #f7f7f7; PADDING-BOTTOM: 4px; MARGIN-LEFT: 10px; BORDER-LEFT: #ccc 1px solid; MARGIN-RIGHT: 20px; PADDING-TOP: 4px; BORDER-BOTTOM: #ccc 1px solid;width:200px;height:20px;
}

.sbut {
	BORDER-RIGHT: #006633 1px solid; BORDER-TOP: #006633 1px solid;  FONT-WEIGHT: bold; FONT-SIZE: 16px; BACKGROUND: url(images/top_bg_hr.jpg) #9c0 repeat-x 1px -10px;  BORDER-LEFT: #006633 1px solid; WIDTH: 88px; COLOR: #ffc; BORDER-BOTTOM: #006633 1px solid; LETTER-SPACING: 4px; HEIGHT: 31px
}
-->
</style></head>

<body>
<table width="664" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="100">&nbsp;</td>
  </tr>
  <tr>
    <td><table width="657" border="0" cellspacing="0" cellpadding="0">

      <tr>
        <td height="169" align="center"><table width="620" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="100" height="50" align="left" valign="middle">&nbsp;</td>
              <td width="370" height="50" align="left" valign="middle" class="t1">管理登录</td>
            </tr>
            <tr>
              <td height="2" colspan="2" align="right" valign="middle" bgcolor="#dddddd"></td>
            </tr>
            <tr>
              <td height="90" colspan="2" align="left" valign="middle"><table width="350" border="0" cellspacing="0" cellpadding="0">
                  <form id="form1" name="form1" method="post" action="?action=1">
                    <tr>
                      <td height="30">&nbsp;</td>
                      <td height="30" colspan="2" align="left" valign="middle">&nbsp;</td>
                    </tr>
                    <tr>
                      <td width="30">&nbsp;</td>
                      <td colspan="2" align="left" valign="middle"><table width="301" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td width="60" height="35" align="center" valign="middle"><span class="t2">用户名：</span></td>
                            <td width="241" height="35" align="left" valign="middle"><input name="name" type="text" class="main" id="name" size="18" maxlength="20" /></td>
                          </tr>
                          <tr>
                            <td height="35" align="center" valign="middle" class="t2">密&nbsp;&nbsp;码：</td>
                            <td height="35" align="left" valign="middle"><input name="pass" type="password" class="main" id="pass" size="18" maxlength="20" /></td>
                          </tr>
                          <tr>
                            <td height="50" colspan="2" align="center" valign="bottom"><input name="Submit" type="submit" class="sbut" value="登 录" /></td>
                            </tr>
                      </table></td>
                      </tr>
                  </form>
              </table></td>
            </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
</body>
</html>
