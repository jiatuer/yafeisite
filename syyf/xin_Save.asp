<!--#include file="include/conn.asp"-->
<!--#include file="Inc/xin_Config.asp"-->
<%
  Session.timeout = 1
  If Config_Web_safe Then
    From_url = Cstr(Request.ServerVariables("HTTP_REFERER"))
    Serv_url = Cstr(Request.ServerVariables("SERVER_NAME"))
    If mid(From_url,8,len(Serv_url)) <> Serv_url then
       Response.Write("<meta http-equiv=""refresh"" content=""2;URL="&config_web_url&""">") 
       Response.Write("�Բ���,��ϵͳ�����˷��ⲿ�ύ����,�벻Ҫ���ⲿ�ύ����,2��󷵻�...") 
       Response.End
    End if
  End if
  If Config_Web_Refresh Then
    Dim URL 
    If DateDiff("s",Request.Cookies("oesun")("vitistime"),Now())<60 Then 
       URL=Request.ServerVariables("Http_REFERER") 
       Response.Write("<meta http-equiv=""refresh"" content=""2;URL="&URL&""">") 
       Response.Write("�Բ���,��ϵͳ�����˷�ˢ�»���,2��󷵻�...") 
       Response.End
    End IF 
    Response.Cookies("oesun")("vitistime")=Now()
  End if
  If Config_Web_Time Then
    IF Session("Send") = 1 Then
      Response.write"<script language=JavaScript>{window.alert('������Ϣ�Ѿ��ύ,�벻Ҫ�����ύ,лл����');window.location.href='javascript:history.go(-1)'}</script>"
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
     Set jmail = Server.CreateObject("JMAIL.Message")   '���������ʼ��Ķ���
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
     'jmail.silent = true    '����������󣬷���FALSE��TRUE��ֵj
     jmail.Charset = "GB2312"     '�ʼ������ֱ���
     jmail.ContentType = "text/html"    '�ʼ��ĸ�ʽΪHTML��ʽ���ı�
     jmail.AddRecipient config_mail_shou     '�ʼ��ռ��˵ĵ�ַ
     jmail.From = config_mail_fa   '�����˵�E-MAIL��ַ
     jmail.MailServerUserName = config_mail_user     '��¼�ʼ����������û��� (�����ʼ���ַ)
     jmail.MailServerPassword = config_mail_pass     '��¼�ʼ������������� (�����ʼ�����)
     jmail.Subject = "��Ϣ����������"    '�ʼ��ı��� 
     jmail.Body = "<font color=#FF0000>����</font>:"&Mail_title&"<br><font color=#FF0000>����</font>:"&Mail_Type&"<br><font color=#FF0000>����</font>:"&Mail_Name&"<br><font color=#FF0000>E-Mail</font>:"&Mail_Mail&"<br><font color=#FF0000>����</font>:"&Mail_From&"<br><font color=#FF0000>��ϵ��ַ</font>:"&Mail_Address&"<br><font color=#FF0000>��ϵ�绰</font>:"&Mail_Phone&"<br><font color=#FF0000>������Ϣ</font>:"&Mail_Info&"<br><font color=#FF0000>����ʱ��</font>:"&Mail_Datetime&"<br><font color=#FF0000>����IP</font>:"&Mail_IP&""      '�ʼ�������
     jmail.Priority = config_mail_cd      '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
     jmail.Send(config_mail_smtp)     'ִ���ʼ����ͣ�ͨ���ʼ���������ַ��
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
      <TD height=298 align="right" bgcolor="#333333"><p align="center">������Ϣ�Ѿ�������������,���ǻᾡ�촦��.лл����֧��!</p>
          <p align="center"><a href="default.asp">������ҳ</a>&nbsp; <a href="javascript:window.close()"> �رմ���</a></TD>
    </TR>
  </TBODY>
</TABLE>
</BODY></HTML>