<%
 Set Config_rs = Server.CreateObject("ADODB.RecordSet") 
 Sql = "select * from xin_Inconfig"
 Config_rs.Open sql,conn,1,3
  config_web_name = config_rs("config_web_name") 
  config_web_url = config_rs("config_web_url")               
  config_web_copy = config_rs("config_web_copy")                 
  config_web_fangshi = config_rs("config_web_fangshi")         
  config_web_ver = config_rs("config_web_ver")    
  config_web_refresh = config_rs("config_web_refresh")
  config_web_time = config_rs("config_web_time")
  config_web_gg = config_rs("config_web_gg")
  config_web_safe = config_rs("config_web_safe")
  config_mail_shou = config_rs("config_mail_shou")   
  config_mail_fa = config_rs("config_mail_fa")
  config_mail_user = config_rs("config_mail_user")
  config_mail_pass = config_rs("config_mail_pass")
  config_mail_smtp = config_rs("config_mail_smtp")
  config_mail_cd = config_rs("config_mail_cd")  
 Config_rs.close
 Set Config_rs=Nothing
%>
<%
  Function iif(e)
    If e = True Then
      iif = "checked"
    End if
  End Function
%>