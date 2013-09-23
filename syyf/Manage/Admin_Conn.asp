<%
 set conn=Server.CreateObject("ADODB.Connection")
 dbpath=Server.MapPath("../DataBase/xinhua.asp")
 conn.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ="&dbpath
%>