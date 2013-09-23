<!-- #include file="Include/csscon.asp" -->
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
<title></title>
<%
if Request("action")=1 then
title=R(Trim(Request("title")))
name=MD5(R(Trim(Request("name"))))

if title="" or name="" then
	Response.Write("<script>alert('请添加CSS样式！');history.go(-1);</script>")
	Response.End()
end if

sql="select * from zm_info where title='"&title&"' and name='"&name&"'"
set rs=server.CreateObject("ADODB.Recordset")
rs.open sql,conn,1,3
if rs.eof and rs.bof then
	Response.Write("<script>alert('CSS样式出错！');document.location.href('news_info.asp');</script>")
	Response.End()
else
	Session("Admin")=rs("name")
	Response.Redirect("cskz.asp")
	
end if

end if
%>
</head>

<body>
<table width="500" border="0" cellspacing="0" cellpadding="0">
  <form id="form1" name="form1" method="post" action="?action=1">
    <tr>
      <td height="30" colspan="3"><table width="100%" border="0" cellpadding="0" cellspacing="0" class="ta">
        <tr>
          <td width="50%" height="25" align="left" valign="middle" class="t2">样式管理</td>
          <td width="50%" align="right" valign="middle" class="t2">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr>
      <td width="30">&nbsp;</td>
      <td colspan="2" align="left" valign="middle"><table width="500" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="35" align="center" valign="middle" class="t2"><span class="t3">CSS样式A：</span></td>
          <td width="241" height="35" align="left" valign="middle"><input name="title" type="text" class="main" id="name" size="18" value="<%=c_name%>" maxlength="20" /></td>
        </tr>
        <tr>
          <td height="35" align="center" valign="middle" class="t2"><span class="t3">CSS样式B：</span></td>
          <td height="35" align="left" valign="middle"><input name="name" type="text" class="main" id="pass" size="18" value="<%=c_adr%>" maxlength="20" /></td>
        </tr>
        <tr>
          <td height="35" align="center" valign="middle" class="t2"><span class="t3">CSS样式C：</span></td>
          <td height="35" align="left" valign="middle"><input name="tel" type="text" class="main" id="pass" size="18" value="<%=c_tel%>"  maxlength="20" /></td>
        </tr>
        <tr>
          <td height="35" align="center" valign="middle" class="t2"><span class="t3">CSS样式D：</span></td>
          <td height="35" align="left" valign="middle"><input name="key" type="text" class="main" id="pass" size="18" value="<%=c_key%>"  maxlength="20" /></td>
        </tr>
        <tr>
          <td height="35" align="center" valign="middle" class="t2"><span class="t3">CSS样式E：</span></td>
          <td height="35" align="left" valign="middle"><input name="fax" type="text" class="main" id="pass" size="18" value="<%=c_fax%>"  maxlength="20" /></td>
        </tr>
        <tr>
          <td height="50" colspan="2" align="center" valign="bottom"><input name="Submit" type="submit" class="sbut" value="提交" /></td>
        </tr>
      </table></td>
    </tr>
  </form>
</table>
</body>
</html>