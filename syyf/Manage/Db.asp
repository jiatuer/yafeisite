<!-- #include file="../Include/conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<%
if Request.QueryString("action")="1" then
	dbpath=Server.MapPath(datapath)
	bakname=year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&".mdb"
	bakpath=Server.MapPath("./dbbak/")
	
	Set Fso=server.createobject("scripting.filesystemobject") 
	fso.copyfile dbpath,bakpath& "\"& bakname
	
	Response.Write("<script>alert('数据库备份成功!');</script>")
end if
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
    <td height="33" background="Images/l_mu_bg2.gif"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="117" height="33" background="Images/l_mu_bg1.gif" class="tt" style="background-repeat:no-repeat">备份数据</td>
        <td width="646" height="33">&nbsp;</td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="10"></td>
  </tr>
  <tr>
    <td><table width="100%" border="0" cellpadding="0" cellspacing="1" class="ta">
      <tr>
        <td height="50" align="center" valign="middle" class="t2"><form name="form1" method="post" action="?action=1" style="margin:0">
          <input name="Submit" type="submit" class="inputbut" value="立即备份">
        </form>
        </td>
        </tr>
      <tr>
        <td height="25" align="center" valign="middle" class="tb">说明：请单击“立即备份”按钮，系统会自动在管理目录下的DBBAK目录下创建一个以当前时间命名的数据库文件。</td>
      </tr>
      
    </table>
	</td>
  </tr>
</table>
</body>
</html>
