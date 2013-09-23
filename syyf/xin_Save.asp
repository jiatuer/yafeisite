<!--#include file="include/conn.asp"-->
<!--#include file="Inc/xin_Config.asp"-->
<%
  Session.timeout = 1
  If Config_Web_safe Then
    From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
    If mid(From_url,8,len(Serv_url)) <> Serv_url then
       Response.Write("<meta http-equiv=""refresh"" content=""2;URL="&config_web_url&""">") 
       Response.Write("对不起,本系统采用了防外部提交机制,请不要从外部提交数据,2秒后返回...") 
       Response.End
    End if
  End if
  If Config_Web_Refresh Then
    Dim URL 
    If DateDiff("s",Request.Cookies("oesun")("vitistime"),Now())<60 Then 
       URL=Request.ServerVariables("Http_REFERER") 
       Response.Write("<meta http-equiv=""refresh"" content=""2;URL="&URL&""">") 
       Response.Write("对不起,本系统采用了防刷新机制,2秒后返回...") 
       Response.End
    End IF 
    Response.Cookies("oesun")("vitistime")=Now()
  End if
  If Config_Web_Time Then
    IF Session("Send") = 1 Then
      Response.write"<script language=JavaScript>{window.alert('您的信息已经提交,请不要继续提交,谢谢合作');window.location.href='javascript:history.go(-1)'}</script>"
    End if
  End if
%>
<%
 Send_id = Request("Action")
 Select Case Send_id
   Case "send"
     Set rs=Server.CreateObject("ADODB.Recordset")
     Sql="Select * from xin_Feedback"
     Rs.Open sql,conn,1,3
     Rs.Addnew
     Rs("Feedback_Title") = Request("Feedback_Title")
     Rs("Feedback_Type") = Request("Feedback_Type")
     Rs("Feedback_Name") = Request("Feedback_Name")
     Rs("Feedback_Qq") = Request("Feedback_Qq")
     Rs("Feedback_Mail") = Request("Feedback_Mail")
     Rs("Feedback_From") = Request("Feedback_From")
     Rs("Feedback_Address") = Request("Feedback_Address")
     Rs("Feedback_Phone") = Request("Feedback_Phone")
     Rs("Feedback_Info") = Request("Feedback_Info")
     Rs("Feedback_zt") = 0
     Rs("Feedback_datetime")= now()
     Rs("Feedback_ip") = Request.ServerVariables("REMOTE_ADDR")
     Rs("Feedback_fs") = "Send"
     Rs.Update
     Rs.Close
     Set rs=Nothing
     Session("Send") = 1
   Case "mail"
     Set jmail = Server.CreateObject("JMAIL.Message")   '建立发送邮件的对象
     Mail_Title = Request("Feedback_Title")
     Mail_Type = Request("Feedback_Type")
     Mail_Name = Request("Feedback_Name")
     Mail_Mail = Request("Feedback_Mail")
     Mail_From = Request("Feedback_From")
     Mail_Address = Request("Feedback_Address")
     Mail_Phone = Request("Feedback_Phone")
     Mail_Info = Request("Feedback_Info")
     Mail_Datetime= now()
     Mail_Ip = Request.ServerVariables("REMOTE_ADDR")     
     'jmail.silent = true    '屏蔽例外错误，返回FALSE跟TRUE两值j
     jmail.Charset = "GB2312"     '邮件的文字编码
     jmail.ContentType = "text/html"    '邮件的格式为HTML格式或纯文本
     jmail.AddRecipient config_mail_shou     '邮件收件人的地址
     jmail.From = config_mail_fa   '发件人的E-MAIL地址
     jmail.MailServerUserName = config_mail_user     '登录邮件服务器的用户名 (您的邮件地址)
     jmail.MailServerPassword = config_mail_pass     '登录邮件服务器的密码 (您的邮件密码)
     jmail.Subject = "信息反馈表单内容"    '邮件的标题 
     jmail.Body = "<font color=#FF0000>标题</font>:"&Mail_title&"<br><font color=#FF0000>类型</font>:"&Mail_Type&"<br><font color=#FF0000>姓名</font>:"&Mail_Name&"<br><font color=#FF0000>E-Mail</font>:"&Mail_Mail&"<br><font color=#FF0000>来自</font>:"&Mail_From&"<br><font color=#FF0000>联系地址</font>:"&Mail_Address&"<br><font color=#FF0000>联系电话</font>:"&Mail_Phone&"<br><font color=#FF0000>反馈信息</font>:"&Mail_Info&"<br><font color=#FF0000>反馈时间</font>:"&Mail_Datetime&"<br><font color=#FF0000>反馈IP</font>:"&Mail_IP&""      '邮件的内容
     jmail.Priority = config_mail_cd      '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
     jmail.Send(config_mail_smtp)     '执行邮件发送（通过邮件服务器地址）
     jmail.Close()   
     set jmail = nothing
     Set rs=Server.CreateObject("ADODB.Recordset")
     Sql="Select * from xin_Feedback"
     Rs.Open sql,conn,1,3
     Rs.Addnew
     Rs("Feedback_Title") = Request("Feedback_Title")
     Rs("Feedback_Type") = Request("Feedback_Type")
     Rs("Feedback_Name") = Request("Feedback_Name")
     Rs("Feedback_Qq") = Request("Feedback_Qq")
     Rs("Feedback_Mail") = Request("Feedback_Mail")
     Rs("Feedback_From") = Request("Feedback_From")
     Rs("Feedback_Address") = Request("Feedback_Address")
     Rs("Feedback_Phone") = Request("Feedback_Phone")
     Rs("Feedback_Info") = Request("Feedback_Info")
     Rs("Feedback_zt") = 1
     Rs("Feedback_datetime")= now()
     Rs("Feedback_ip") = Request.ServerVariables("REMOTE_ADDR")
     Rs("Feedback_fs") = "Mail"
     Rs("Feedback_yj") = Config_Mail_Shou
     Rs.update
     Rs.close     
     Set rs = nothing
     Session("Send") = 1
   End Select
%>
<HTML>
<HEAD>
<TITLE><%=config_web_name%></TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<STYLE type=text/css>.MainMenuBGStyle {
	BACKGROUND-POSITION: right center; BACKGROUND-REPEAT: no-repeat
}
</STYLE>

<style type="text/css">
<!--
body,td,th {
	color: #FFFFFF;
}
body {
	background-color: #000000;
}
a:link {
	color: #FFFFFF;
}
a:visited {
	color: #FFFFFF;
}
-->
</style></head>
<TABLE width=699 border=0 align="center" cellPadding=5 cellSpacing=1 
      bgColor=#333333>
  <TBODY>
    <TR bgColor=#fafafa>
      <TD height=298 align="right" bgcolor="#333333"><p align="center">您的信息已经反馈给了我们,我们会尽快处理.谢谢您的支持!</p>
          <p align="center"><a href="default.asp">返回首页</a>&nbsp; <a href="javascript:window.close()"> 关闭窗口</a></TD>
    </TR>
  </TBODY>
</TABLE>
</BODY></HTML>