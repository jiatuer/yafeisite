<!-- #include file="../Include/conn.asp" -->
<!--#include file="../Inc/xin_Config.asp" -->
<!-- #include file="Include/Chk.asp" -->
<%
 id = request("id")
 Set rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Feedback Where Feedback_id ="&id
 Rs.Open sql,conn,1,3
 If rs.eof and rs.bof Then
   Response.write ("没有你要找的记录")
 Else
   If Rs("Feedback_zt") <> 1 Then
     Rs("Feedback_zt") = 1
   End if
   Rs.update
   Rs.close
 End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=config_web_name%></title>
<STYLE type=text/css>.MainMenuBGStyle {
	BACKGROUND-POSITION: right center; BACKGROUND-REPEAT: no-repeat
}
</STYLE>
<LINK href="Images/infocss.css" rel=stylesheet>
</head>
<body topmargin="1" leftmargin="1" rightmargin="1" bottommargin="1" marginwidth="1" marginheight="1">
<div align=center>
      <TABLE height=30 cellSpacing=0 cellPadding=0 width=446 
      background=Images/main_bg.gif border=0>
        <TBODY>
        <TR>
          <TD>　<strong><font color="#006699">反馈信息管理</font></strong></TD></TR></TBODY></TABLE>
      <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
        <TBODY>
        <TR>
          <TD height=1></TD></TR></TBODY></TABLE>
      <div align="center">
      <TABLE cellSpacing=1 cellPadding=5 width=446 
      bgColor=#C0C0C0 border=0 height="286">
        <TBODY>
        <TR bgColor=#fafafa>
          <TD height=284>
          <div align=center>
            <table border="0" width="428" id="table4" height="300">
<%
 id = request("id")
 Set rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Feedback Where Feedback_id ="&id
 Rs.Open sql,conn,1,3
 If not rs.eof then
%>
				<tr>
					<td bgcolor="#CDCDCD" align="right" width="112" height="19">
					<p align="right"><b>序号:</b></td>
					<td bgcolor="#F0F0F0" align="left" width="306" height="22"><%=Rs("Feedback_id")%></td>
				</tr>
				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>房型或建筑面积:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22">
					<font color="#FF0000"><%=Rs("Feedback_title")%></font></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>类型:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22"><%=Rs("Feedback_type")%></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>姓名:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22"><%=Rs("Feedback_name")%></td>
				</tr>

<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>Q Q:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22"><%=Rs("Feedback_qq")%></td>
				</tr>
				
				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>E-Mail:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22"><%=Rs("Feedback_mail")%></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>房屋所属区:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22"><%=Rs("Feedback_from")%></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>联系地址:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22"><%=Rs("Feedback_address")%></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>电话:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="22"><%=Rs("Feedback_phone")%></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="63">
					<b>附加要求和说明:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="center" height="63">
					<p align="left">
					<textarea rows="6" name="S1" cols="38"><%=Rs("Feedback_info")%></textarea></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>反馈日期:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="21"><%=Rs("Feedback_datetime")%></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>反馈IP:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="21"><%=Rs("Feedback_ip")%></td>
				</tr>

				<tr>
					<td width="112" bgcolor="#CDCDCD" align="right" height="22">
					<b>当前状态:</b></td>
					<td width="306" bgcolor="#F0F0F0" align="left" height="10"><% Select Case Rs("Feedback_zt")
					Case 0
					  response.write "<font color=#FF0000><b>未读</b></font>"
					Case 1
					  response.write "<b>已读</b>"
					End Select%></td>
				</tr>
<tr>
					<td width="418" bgcolor="#E6E6E6" align="right" height="28" colspan="2">
					<p align="center"><a href="#" onClick="javascript:window.close();">关闭窗口</a></td>
					</tr>
<%
 Else
%>
				<tr>
					<td width="422" bgcolor="#E6E6E6" align="right" height="26" colspan="2">
					<p align="center"><font color="#FF0000"><b>对不起,没有找到您需要的记录</b></font></td>
				</tr>
<%
 End if
%>
			</table>
<%
 Rs.close
 Set rs = nothing
%>
			</div>			</TD>
          </TR></TBODY></TABLE>
</div>            
</body>

</html>