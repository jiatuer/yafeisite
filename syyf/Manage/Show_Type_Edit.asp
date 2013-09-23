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
    <td><table width="100%" border="0" cellpadding="0" cellspacing="0" class="ta">
      <tr>
        <td width="50%" height="25" align="left" valign="middle" class="t2">修改分类</td>
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
		  <form name="form" method="post" action="?action=1&id=<%=request("id")%>">
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">分类名称：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="titles" type="text" class="main" id="name" value="" size="20" maxlength="20"></td>
            </tr>
            <tr>
              <td width="40%" height="30" align="right" valign="middle" class="t3">显示顺序：</td>
              <td width="60%" height="30" align="left" valign="middle"><input name="ord" type="text" class="main" id="ord" value="" size="4" maxlength="4">
                <span class="ts">*数值越小越靠前</span></td>
            </tr>
            <tr>
              <td height="50" colspan="2" align="center" valign="middle"><a href="Admin_xin.asp"><img src="Images/xin1.jpg" width="140" height="38" border="0"></a></td>
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