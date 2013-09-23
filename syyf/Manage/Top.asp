<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title></title>

<link href="Images/css.css" rel="stylesheet" type="text/css" />
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="300" height="74" align="left" background="Images/top_bg.gif"><img src="Images/logo2.gif" width="300" height="74" /></td>
    <td height="74" align="right" valign="bottom" background="Images/top_bg.gif"><table width="390" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="150" height="40" align="center" valign="middle" class="tfont">欢迎你，<%=session("admin")%>！</td>
        <td width="80" align="center" valign="middle" class="tfont"><a href="main.asp" target="main">管理首页</a></td>
        <td width="80" align="center" valign="middle" class="tfont"><a href="/" target="_blank">网站首页</a></td>
        <td width="80" height="40" align="center" valign="middle" class="tfont"><a href="Exit.asp" target="_parent" ONCLICK="javascript:return confirm('提示：您确定要退出吗？')">退出系统</a></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
</html>
