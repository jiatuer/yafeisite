<!--#include file="include/conn.asp"-->
<!--#include file="Inc/xin_Config.asp"-->
<HTML>
<HEAD>
<TITLE>在线装修</TITLE>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<link href="images/xin_main.css" rel="stylesheet" type="text/css">
<script language="VBscript">
<!--
Sub checkdata()
	If Trim(YuQaIFS_Feedback.Feedback_Type.value)=empty Then
		Window.Alert"类型不能为空"
		Exit Sub
	End If
	If Trim(YuQaIFS_Feedback.Feedback_Name.value)=empty Then
		Window.Alert"姓名不能为空"
		Exit Sub
	End If
        If Trim(YuQaIFS_Feedback.Feedback_qq.value)=empty Then
		Window.Alert"QQ不能为空"
		Exit Sub
	End If
	If Trim(YuQaIFS_Feedback.Feedback_From.value)=empty Then
		Window.Alert"来自地区不能为空"
		Exit Sub
	End If
		If Trim(YuQaIFS_Feedback.Feedback_Phone.value)=empty Then
		Window.Alert"电话不能为空"
		Exit Sub
	End If
	If Trim(YuQaIFS_Feedback.Feedback_Info.value)=empty Then
		Window.Alert"反馈信息不能为空"
		Exit Sub
	End If
YuQaIFS_Feedback.submit	
End Sub
-->
</script>
<SCRIPT src="images/jquery.js" type=text/javascript></SCRIPT>
<SCRIPT src="images/FunNav.js" type=text/javascript></SCRIPT>

<style type="text/css">
<!--
#Layer1 {	position:absolute;
	width:235px;
	height:120px;
	z-index:1;
	top: 33px;
	overflow: visible;
}
-->
</style>
</head>
<body>
<table width="999" height="159" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td background="images/headerbj.jpg"><table width="1000" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="32">&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td width="230" height="100"><div align="center">
          <div id="Layer1">
            <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,19,0" width="229" height="109">
              <param name="movie" value="Images/xi.swf" />
              <param name="quality" value="high" />
              <param name="wmode" value="transparent" />
              <embed src="Images/xi.swf" width="229" height="109" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" wmode="transparent"></embed>
            </object>
          </div>
          <img src="images/LOGO.GIF" width="208" height="81" /></div></td>
        <td><table width="742" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td height="50"><div align="right"><img src="images/tel.png" width="260" height="25" /></div></td>
          </tr>
          <tr>
            <td height="42"><table width="736" border="0" align="center" cellpadding="0" cellspacing="0">
              <tr>
                <td><a href="home.asp"><img src="images/top0.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top00.jpg';" onMouseOut="this.src='images/top0.jpg';"/></a></td>
                <td><a href="about.asp?id=17"><img src="Images/top1.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top01.jpg';" onMouseOut="this.src='images/top1.jpg';"/></a></td>
                <td><a href="news.asp"><img src="Images/top2.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top02.jpg';" onMouseOut="this.src='images/top2.jpg';"/></a></td>
                <td><a href="show.asp"><img src="Images/top3.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top03.jpg';" onMouseOut="this.src='images/top3.jpg';"/></a></td>
                <td><a href="product.asp"><img src="Images/top4.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top04.jpg';" onMouseOut="this.src='images/top4.jpg';"/></a></td>
                <td><a href="message.asp"><img src="Images/top5.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top05.jpg';" onMouseOut="this.src='images/top5.jpg';"/></a></td>
                <td><a href="service.asp"><img src="Images/top6.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top06.jpg';" onMouseOut="this.src='images/top6.jpg';"/></a></td>
                <td><a href="about.asp?id=21"><img src="Images/top7.jpg" width="92" height="40" border="0"  onmousemove="this.src='images/top07.jpg';" onMouseOut="this.src='images/top7.jpg';"/></a></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="9"></td>
          </tr>
      </table></td>
  </tr>
</table>
<table width="999" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><img src="images/fbanner.jpg" width="998" height="90" /></td>
  </tr>
</table>
<table width="999" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000">
  <tr>
    <td height="10"></td>
  </tr>
</table>
<TABLE width="999" border=0 align="center" cellPadding=0 cellSpacing=0 id=nav_main>
  <TBODY>
  <TR>
    <TD width="211" valign="top" background="images/mainflbj.png"><table width="150" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="190" valign="top">&nbsp;</td>
      </tr>
    </table>
      <table width="211" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><img src="images/footfbj1-1.png" width="211" height="89" /></td>
        </tr>
        <tr>
          <td><img src="images/footfbj1-2.png" width="211" height="89" /></td>
        </tr>
        <tr>
          <td><img src="images/footfbj1-3.png" width="211" height="64" /></td>
        </tr>
        <tr>
          <td><img src="images/footfbj2.png" width="211" height="49" /></td>
        </tr>
        <tr>
          <td><img src="images/footfbj3.png" width="211" height="47" /></td>
        </tr>
      </table></TD>
    <TD width="618" height="500" valign="top" bgcolor="#333333"><table width="100%" height="100" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><div align="center"><img src="images/Service.jpg" width="561" height="91"></div></td>
      </tr>
    </table>
      <TABLE cellSpacing=1 cellPadding=5 width=100% 
      bgColor=#cdcdcd border=0>
        <form method="POST" action="xin_Save.asp?Action=<%=config_web_fangshi%>" name="YuQaIFS_Feedback">
          <TBODY>
            <TR bgColor=#fafafa>
              <TD height=28 colspan="3" align="right" bgcolor="#333333"><div align="left"><br>
                &nbsp;以下带&quot;<font color="#FF0000">*</font>&quot;部分必须填写 </div></TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD width="100" height=28 align="right" bgcolor="#333333"> 类&nbsp; 型:</TD>
              <TD width="300" height=28 bgcolor="#333333"><select name="Feedback_Type">
                  <option selected>请选择</option>
                  <%
 Set rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Types"
 Rs.Open sql,conn,1,3
 If not rs.eof then
   Count = rs.recordcount
   For i = 1 to count
     response.write "<option value='"&rs("Type_Value")&"'>"&rs("Type_Value")&"</option>"
     rs.movenext
   Next
 End if
 rs.close
 Set rs=Nothing 
%>
              </select></TD>
              <TD width="30" bgcolor="#333333"><font color="#FF0000">*</font></TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD width="100" height=28 align="right" bgcolor="#333333"> 姓&nbsp; 名:</TD>
              <TD height=28 bgcolor="#333333"><input type="text" maxlength=50 name="Feedback_Name" size="16"></TD>
              <TD width="30" bgcolor="#333333"><font color="#FF0000">*</font></TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD width="100" height=28 align="right" bgcolor="#333333"><div align="right">E-Mail:</div></TD>
              <TD height=28 bgcolor="#333333"><input type="text" maxlength=50 name="Feedback_Mail" size="22"></TD>
              <TD width="30" bgcolor="#333333">&nbsp;</TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD width="100" height=28 align="right" bgcolor="#333333"><div align="right">Q Q:</div></TD>
              <TD height=28 bgcolor="#333333"><input type="text" maxlength=50 name="Feedback_qq" size="22"></TD>
              <TD width="30" bgcolor="#333333"><font color="#FF0000">*</font></TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD height=28 align="right" bgcolor="#333333">房型或建筑面积:</TD>
              <TD height=28 bgcolor="#333333"><input type="text" name="Feedback_Title" maxlength=50 size="41"></TD>
              <TD bgcolor="#333333">&nbsp;</TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD height=28 align="right" bgcolor="#333333">房屋所属区:</TD>
              <TD height=28 bgcolor="#333333"><select name="Feedback_From">
                  <option selected>请选择</option>
                  <%
 Set rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Froms"
 Rs.Open sql,conn,1,3
 If not rs.eof then
   Count = rs.recordcount
   For i = 1 to count
     response.write "<option value='"&rs("From_Value")&"'>"&rs("From_Value")&"</option>"
     rs.movenext
   Next
 End if
 rs.close
 Set rs=Nothing 
%>
              </select></TD>
              <TD bgcolor="#333333"><font color="#FF0000">*</font></TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD height=28 align="right" bgcolor="#333333"> 联系地址:</TD>
              <TD height=28 bgcolor="#333333"><input type="text" maxlength=50 name="Feedback_Address" size="35"></TD>
              <TD bgcolor="#333333">　</TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD height=27 align="right" bgcolor="#333333"> 电&nbsp; 话:</TD>
              <TD height=27 bgcolor="#333333"><input type="text" maxlength=50 name="Feedback_Phone" size="19"></TD>
              <TD bgcolor="#333333"><font color="#FF0000">*</font></TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD height=127 align="right" bgcolor="#333333">附加要求和说明:</TD>
              <TD height=127 bgcolor="#333333"><textarea rows="9" name="Feedback_Info" cols="43" ></textarea></TD>
              <TD bgcolor="#333333"><font color="#FF0000">*</font></TD>
            </TR>
            <TR bgColor=#fafafa>
              <TD height=28 colspan="3" align="right" bgcolor="#333333"><p align="center">
                  <input type="button" onclick=checkdata() value="提 交 信 息" name="tijiao" class="input2">
                  <span lang="zh-cn">&nbsp;&nbsp;&nbsp;&nbsp; </span>&nbsp;
                  <input type="reset" value="重 新 填 写" name="reset" class="input2">
                </TD>
            </TR>
          </TBODY>
        </form>
      </TABLE></TD>
    <TD width="170" bgcolor="#000000" id=nav_mainr>
      <DIV id=navfuntop></DIV>
      <DIV id=navfun></DIV></TD></TR>
  <TR>
    <TD background="images/mainflbj.png">&nbsp; </TD>
    <TD bgcolor="#333333">&nbsp; </TD>
<TD bgcolor="#000000">&nbsp; </TD>
  </TR></TBODY></TABLE>

<table width="999" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#000000">
  <tr>
    <td><DIV id=footerf></DIV></td>
  </tr>
</table>
</BODY></HTML>