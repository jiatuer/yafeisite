<!-- #include file=aasql.asp -->
<!-- #include file=function.asp -->
<%
'On Error Resume Next
   dim connstr,datapath,conn
   datapath="../tools/css.mdb"
   connstr="Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & Server.mappath(datapath)
   Set conn=Server.CreateObject("ADODB.Connection")
   conn.open connstr

%>