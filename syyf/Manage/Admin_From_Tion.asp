<!-- #include file="Admin_Conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!--#include file="../Inc/xin_Config.asp" -->
<!--#include file="../Inc/xin_Md5.asp" -->
<%
  Action = request("Action")
  Select Case Action
    Case 1                         ' ��Ϣɾ��
      Action_id = request("id")
      Action_zt = request("zt")
      Action_fs = request("fs")
      If Action_zt = "w" Then
        response.write"<script language=JavaScript>{window.alert('������Ϣ����δ��!');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Set rs=Server.CreateObject("ADODB.Recordset")      
        Sql="Delete * from xin_Feedback where Feedback_id="&Action_id
        Rs.Open Sql,conn,1,3
        if Action_fs = "mail" Then
          response.write"<script language=JavaScript>{window.alert('���Ѿ��ɹ�ɾ��һ���ʼ���Ϣ');window.location.href='Admin_Mail.asp'}</script>"      
        Else
          response.write"<script language=JavaScript>{window.alert('���Ѿ��ɹ�ɾ��һ����̨��Ϣ');window.location.href='Admin_Msg.asp'}</script>"
        End if
        Set Rs = nothing
      End if
    Case 2                         ' �����˵����
      Title = request("Title")
      If Len(Title) > 10 or Title = "" or Title = Null Then
        response.write"<script language=JavaScript>{window.alert('�Բ���,��������ַ������Ϲ淶');window.location.href='javascript:history.go(-1)'}</script>"
      End if
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Froms where From_value = '"&Title&"'"
      Rs.open sql,conn,1,3
      If not Rs.eof then
        Response.write"<script language=JavaScript>{window.alert('�Բ���,�������˵������Ѿ�����,����������');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs.addnew
        Rs("From_value")=trim(Title)
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('���ɹ��������һ�������˵�');window.location.href='Admin_from.asp?action=1'}</script>"
      End if      
    Case 3                         ' �����˵��޸�
      id = request("id")
      Title = request("Title")
      If Len(Title) > 10 or Title = "" or Title = Null Then
        response.write"<script language=JavaScript>{window.alert('�Բ���,��������ַ������Ϲ淶');window.location.href='javascript:history.go(-1)'}</script>"
      End if
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Froms where From_id = "&id
      Rs.open sql,conn,1,3
      If Title = rs("From_value") then
        Response.write"<script language=JavaScript>{window.alert('�Բ���,�������˵������Ѿ�����,����������');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs("From_value")=trim(Title)
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('���ɹ����޸���һ�������˵�');window.location.href='Admin_from.asp?action=1'}</script>"
      End if    
    Case 4                         ' �����˵�ɾ��
      Action_id = request("id")
      Set rs=Server.CreateObject("ADODB.Recordset")      
      Sql="Delete * from xin_Froms where From_id="&Action_id
      Rs.Open Sql,conn,1,1          
      response.write"<script language=JavaScript>{window.alert('���Ѿ��ɹ�ɾ��һ����¼');window.location.href='Admin_from.asp?Action=1'}</script>"
      Set Rs = nothing
    Case 5                         ' �����˵�ȫ��ɾ��
      Set rs=Server.CreateObject("ADODB.Recordset")      
      Sql="Delete * from xin_Froms"
      Rs.Open Sql,conn,1,1          
      response.write"<script language=JavaScript>{window.alert('���Ѿ��ɹ�ȫ����¼');window.location.href='Admin_from.asp?Action=1'}</script>"
      Set rs = nothing  
    Case 6                         ' ������Ϣȫ��ɾ��
      fs = request("fs")
      If fs = "Mail" Then 
        Set rs=Server.CreateObject("ADODB.Recordset")
        Sql="Delete * from xin_Feedback where Feedback_fs = 'Mail'"
        Rs.Open Sql,conn,1,1          
        response.write"<script language=JavaScript>{window.alert('���Ѿ��ɹ����ȫ���ʼ���¼');window.location.href='Admin_Mail.asp'}</script>"
        Set rs = nothing
      Else
        Set rs=Server.CreateObject("ADODB.Recordset")
        Sql="Delete * from xin_Feedback where Feedback_zt = 1"
        Rs.Open Sql,conn,1,1          
        response.write"<script language=JavaScript>{window.alert('���Ѿ��ɹ�����Ѷ���̨��¼');window.location.href='Admin_Msg.asp'}</script>"
        Set rs = nothing
      End if        
    Case 7                         ' ϵͳ����
      name = request("name")
      url = request("url")
      fangshi = request("fangshi")
      refresh = request("refresh")
      qtime = request("time")
      gg = request("gg")
      safe = request("safe")
      Set Rs = Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Inconfig"
      Rs.open sql,conn,1,3
      Rs("Config_web_name")=trim(name)
      Rs("Config_web_url")=trim(url)
      Rs("Config_web_fangshi")=trim(fangshi)
      Rs("Config_web_refresh")=refresh
      Rs("Config_web_time")=qtime
      Rs("Config_web_gg")=gg
      Rs("Config_web_safe")=safe
      Rs.update
      Rs.close     
      Set rs = nothing
      response.write"<script language=JavaScript>{window.alert('�������óɹ�');window.location.href='Admin_InfoConfig.asp'}</script>" 
    Case 8                         ' �ʼ�����
      shou = request("shou")
      fa = request("fa")
      user = request("user")
      pass = request("pass")
      smtp = request("smtp")
      cd = request("cd")
      Set Rs = Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Inconfig"
      Rs.open sql,conn,1,3
      Rs("Config_mail_shou")=trim(shou)
      Rs("Config_mail_fa")=trim(fa)
      Rs("Config_mail_user")=trim(user)
      Rs("Config_mail_pass")=trim(pass)
      Rs("Config_mail_smtp")=trim(smtp)
      Rs("Config_mail_cd")=trim(cd)
      Rs.update
      Rs.close     
      Set rs = nothing
      response.write"<script language=JavaScript>{window.alert('�������óɹ�');window.location.href='Admin_JMail.asp'}</script>" 
    Case 9                         ' ���ӹ���Ա
      User = request("User")
      Pass = request("Pass")
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Admin where Admin_User= '"&User&"'"
      Rs.open sql,conn,1,3
      If not Rs.eof then
        Response.write"<script language=JavaScript>{window.alert('�Բ���,���û����Ѿ�����,����������');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs.addnew
        Rs("Admin_User")=trim(User)
        Rs("Admin_Pass")=Md5(trim(Pass),Pmode)
        Rs("Admin_Lock")="0"
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('���ɹ������һ������Ա');window.location.href='Admin_Admin.asp?action=1'}</script>"
      End if
    Case 10                        ' ����Ա�����޸�
      id = request("id")
      ModiPass = Md5(request("Pass"),Pmode)
      NewPass = Md5(request("NewPass"),Pmode)
      QNewPass = Md5(request("QNewPass"),Pmode)
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Admin where Admin_id = "&id
      Rs.open sql,conn,1,3
      If ModiPass <> Rs("Admin_Pass") Then
        response.write"<script language=JavaScript>{window.alert('�Բ���,���������������');window.location.href='javascript:history.go(-1)'}</script>"
      End if
      If NewPass <> QNewPass then
        Response.write"<script language=JavaScript>{window.alert('�Բ���,��������������벻ͬ,����������');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs("Admin_Pass")=trim(NewPass)
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('���ɹ��޸��˹���Ա����');window.location.href='Admin_Admin.asp?action=1'}</script>"
      End if
    Case 11                        ' ����Աɾ��
      id = request("id")
      Set rs=Server.CreateObject("ADODB.Recordset")      
      Sql="Delete * from xin_Admin where Admin_id="&id
      Rs.Open Sql,conn,1,1          
      response.write"<script language=JavaScript>{window.alert('���Ѿ��ɹ�ɾ��һ����¼');window.location.href='Admin_Admin.asp?Action=1'}</script>"
      Set Rs = nothing
    Case 12                        ' ����Ա״̬�޸�
      id = request("id")
      Lock = request("Lock")
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Admin where Admin_id = "&id
      Rs.open sql,conn,1,3
      Rs("Admin_Lock")=Lock
      Rs.update
      Rs.close     
      Set rs = nothing
      response.write"<script language=JavaScript>{window.alert('���ɹ��Թ���Ա״̬�����˲���');window.location.href='Admin_Admin.asp?action=1'}</script>"
    Case 13                        ' ����ʼ���¼
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Inconfig"
      Rs.open sql,conn,1,3
      Rs("Config_Mail_Count")=0
      Rs.update
      Rs.close     
      Set rs = nothing
      response.write"<script language=JavaScript>{window.alert('����ɹ�');window.location.href='Admin_InfoMain.asp'}</script>"
    Case 14                        ' ������Ϣ�ظ�
      Set jmail = Server.CreateObject("JMAIL.Message")   '���������ʼ��Ķ���
      Mail_Mail = Request("Mail_Mail")
      Mail_Title = Request("Mail_Title")
      Mail_Info = Request("Mail_Info")
      M_Title = Request("M_Title")
      'jmail.silent = true    '����������󣬷���FALSE��TRUE��ֵj
      jmail.Charset = "GB2312"     '�ʼ������ֱ���
      jmail.ContentType = "text/html"    '�ʼ��ĸ�ʽΪHTML��ʽ���ı�
      jmail.AddRecipient Mail_Mail     '�ʼ��ռ��˵ĵ�ַ
      jmail.From = config_mail_fa   '�����˵�E-MAIL��ַ
      jmail.MailServerUserName = config_mail_user     '��¼�ʼ����������û��� (�����ʼ���ַ)
      jmail.MailServerPassword = config_mail_pass     '��¼�ʼ������������� (�����ʼ�����)
      jmail.Subject = Mail_Title    '�ʼ��ı��� 
      jmail.Body = "����[<u>"&config_web_name&"</u>]�����<u><b>"&M_Title&"</b></u>����Ϊ���������´�:<br>"&Mail_Info&""      '�ʼ�������
      jmail.Priority = config_mail_cd      '�ʼ��Ľ�������1 Ϊ��죬5 Ϊ������ 3 ΪĬ��ֵ
      jmail.Send(config_mail_smtp)     'ִ���ʼ����ͣ�ͨ���ʼ���������ַ��
      jmail.Close()   
      set jmail = nothing
      response.write"<script language=JavaScript>{window.alert('�ʼ����ͳɹ�');window.location.href='javascript:history.go(-1)'}</script>"
  End Select
%>