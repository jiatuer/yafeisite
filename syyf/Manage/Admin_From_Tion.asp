<!-- #include file="Admin_Conn.asp" -->
<!-- #include file="Include/Chk.asp" -->
<!--#include file="../Inc/xin_Config.asp" -->
<!--#include file="../Inc/xin_Md5.asp" -->
<%
  Action = request("Action")
  Select Case Action
    Case 1                         ' 信息删除
      Action_id = request("id")
      Action_zt = request("zt")
      Action_fs = request("fs")
      If Action_zt = "w" Then
        response.write"<script language=JavaScript>{window.alert('本条消息您还未读!');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Set rs=Server.CreateObject("ADODB.Recordset")      
        Sql="Delete * from xin_Feedback where Feedback_id="&Action_id
        Rs.Open Sql,conn,1,3
        if Action_fs = "mail" Then
          response.write"<script language=JavaScript>{window.alert('您已经成功删除一条邮件信息');window.location.href='Admin_Mail.asp'}</script>"      
        Else
          response.write"<script language=JavaScript>{window.alert('您已经成功删除一条后台信息');window.location.href='Admin_Msg.asp'}</script>"
        End if
        Set Rs = nothing
      End if
    Case 2                         ' 下拉菜单添加
      Title = request("Title")
      If Len(Title) > 10 or Title = "" or Title = Null Then
        response.write"<script language=JavaScript>{window.alert('对不起,您输入的字符不符合规范');window.location.href='javascript:history.go(-1)'}</script>"
      End if
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Froms where From_value = '"&Title&"'"
      Rs.open sql,conn,1,3
      If not Rs.eof then
        Response.write"<script language=JavaScript>{window.alert('对不起,该下拉菜单名称已经存在,请重新输入');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs.addnew
        Rs("From_value")=trim(Title)
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('您成功的添加了一条下拉菜单');window.location.href='Admin_from.asp?action=1'}</script>"
      End if      
    Case 3                         ' 下拉菜单修改
      id = request("id")
      Title = request("Title")
      If Len(Title) > 10 or Title = "" or Title = Null Then
        response.write"<script language=JavaScript>{window.alert('对不起,您输入的字符不符合规范');window.location.href='javascript:history.go(-1)'}</script>"
      End if
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Froms where From_id = "&id
      Rs.open sql,conn,1,3
      If Title = rs("From_value") then
        Response.write"<script language=JavaScript>{window.alert('对不起,该下拉菜单名称已经存在,请重新输入');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs("From_value")=trim(Title)
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('您成功的修改了一条下拉菜单');window.location.href='Admin_from.asp?action=1'}</script>"
      End if    
    Case 4                         ' 下拉菜单删除
      Action_id = request("id")
      Set rs=Server.CreateObject("ADODB.Recordset")      
      Sql="Delete * from xin_Froms where From_id="&Action_id
      Rs.Open Sql,conn,1,1          
      response.write"<script language=JavaScript>{window.alert('您已经成功删除一条记录');window.location.href='Admin_from.asp?Action=1'}</script>"
      Set Rs = nothing
    Case 5                         ' 下拉菜单全部删除
      Set rs=Server.CreateObject("ADODB.Recordset")      
      Sql="Delete * from xin_Froms"
      Rs.Open Sql,conn,1,1          
      response.write"<script language=JavaScript>{window.alert('您已经成功全部记录');window.location.href='Admin_from.asp?Action=1'}</script>"
      Set rs = nothing  
    Case 6                         ' 反馈信息全部删除
      fs = request("fs")
      If fs = "Mail" Then 
        Set rs=Server.CreateObject("ADODB.Recordset")
        Sql="Delete * from xin_Feedback where Feedback_fs = 'Mail'"
        Rs.Open Sql,conn,1,1          
        response.write"<script language=JavaScript>{window.alert('您已经成功清除全部邮件记录');window.location.href='Admin_Mail.asp'}</script>"
        Set rs = nothing
      Else
        Set rs=Server.CreateObject("ADODB.Recordset")
        Sql="Delete * from xin_Feedback where Feedback_zt = 1"
        Rs.Open Sql,conn,1,1          
        response.write"<script language=JavaScript>{window.alert('您已经成功清除已读后台记录');window.location.href='Admin_Msg.asp'}</script>"
        Set rs = nothing
      End if        
    Case 7                         ' 系统设置
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
      response.write"<script language=JavaScript>{window.alert('参数设置成功');window.location.href='Admin_InfoConfig.asp'}</script>" 
    Case 8                         ' 邮件设置
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
      response.write"<script language=JavaScript>{window.alert('参数设置成功');window.location.href='Admin_JMail.asp'}</script>" 
    Case 9                         ' 增加管理员
      User = request("User")
      Pass = request("Pass")
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Admin where Admin_User= '"&User&"'"
      Rs.open sql,conn,1,3
      If not Rs.eof then
        Response.write"<script language=JavaScript>{window.alert('对不起,该用户名已经存在,请重新输入');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs.addnew
        Rs("Admin_User")=trim(User)
        Rs("Admin_Pass")=Md5(trim(Pass),Pmode)
        Rs("Admin_Lock")="0"
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('您成功添加了一个管理员');window.location.href='Admin_Admin.asp?action=1'}</script>"
      End if
    Case 10                        ' 管理员密码修改
      id = request("id")
      ModiPass = Md5(request("Pass"),Pmode)
      NewPass = Md5(request("NewPass"),Pmode)
      QNewPass = Md5(request("QNewPass"),Pmode)
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Admin where Admin_id = "&id
      Rs.open sql,conn,1,3
      If ModiPass <> Rs("Admin_Pass") Then
        response.write"<script language=JavaScript>{window.alert('对不起,您的密码输入错误');window.location.href='javascript:history.go(-1)'}</script>"
      End if
      If NewPass <> QNewPass then
        Response.write"<script language=JavaScript>{window.alert('对不起,您两次输入的密码不同,请重新输入');window.location.href='javascript:history.go(-1)'}</script>"
      Else
        Rs("Admin_Pass")=trim(NewPass)
        Rs.update
        Rs.close     
        Set rs = nothing
        response.write"<script language=JavaScript>{window.alert('您成功修改了管理员密码');window.location.href='Admin_Admin.asp?action=1'}</script>"
      End if
    Case 11                        ' 管理员删除
      id = request("id")
      Set rs=Server.CreateObject("ADODB.Recordset")      
      Sql="Delete * from xin_Admin where Admin_id="&id
      Rs.Open Sql,conn,1,1          
      response.write"<script language=JavaScript>{window.alert('您已经成功删除一条记录');window.location.href='Admin_Admin.asp?Action=1'}</script>"
      Set Rs = nothing
    Case 12                        ' 管理员状态修改
      id = request("id")
      Lock = request("Lock")
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Admin where Admin_id = "&id
      Rs.open sql,conn,1,3
      Rs("Admin_Lock")=Lock
      Rs.update
      Rs.close     
      Set rs = nothing
      response.write"<script language=JavaScript>{window.alert('您成功对管理员状态进行了操作');window.location.href='Admin_Admin.asp?action=1'}</script>"
    Case 13                        ' 清空邮件记录
      Set Rs=Server.CreateObject("ADODB.Recordset")
      Sql = "select * from xin_Inconfig"
      Rs.open sql,conn,1,3
      Rs("Config_Mail_Count")=0
      Rs.update
      Rs.close     
      Set rs = nothing
      response.write"<script language=JavaScript>{window.alert('清除成功');window.location.href='Admin_InfoMain.asp'}</script>"
    Case 14                        ' 反馈信息回复
      Set jmail = Server.CreateObject("JMAIL.Message")   '建立发送邮件的对象
      Mail_Mail = Request("Mail_Mail")
      Mail_Title = Request("Mail_Title")
      Mail_Info = Request("Mail_Info")
      M_Title = Request("M_Title")
      'jmail.silent = true    '屏蔽例外错误，返回FALSE跟TRUE两值j
      jmail.Charset = "GB2312"     '邮件的文字编码
      jmail.ContentType = "text/html"    '邮件的格式为HTML格式或纯文本
      jmail.AddRecipient Mail_Mail     '邮件收件人的地址
      jmail.From = config_mail_fa   '发件人的E-MAIL地址
      jmail.MailServerUserName = config_mail_user     '登录邮件服务器的用户名 (您的邮件地址)
      jmail.MailServerPassword = config_mail_pass     '登录邮件服务器的密码 (您的邮件密码)
      jmail.Subject = Mail_Title    '邮件的标题 
      jmail.Body = "您在[<u>"&config_web_name&"</u>]提出的<u><b>"&M_Title&"</b></u>我们为您做出如下答复:<br>"&Mail_Info&""      '邮件的内容
      jmail.Priority = config_mail_cd      '邮件的紧急程序，1 为最快，5 为最慢， 3 为默认值
      jmail.Send(config_mail_smtp)     '执行邮件发送（通过邮件服务器地址）
      jmail.Close()   
      set jmail = nothing
      response.write"<script language=JavaScript>{window.alert('邮件发送成功');window.location.href='javascript:history.go(-1)'}</script>"
  End Select
%>