<!-- #include file="include/conn.asp" -->
<!-- #include file="include/vars.asp" -->
<!-- #include file="include/asptemplate.asp" -->
<!-- #include file="xin_main.asp" -->
<!-- #include file="xin_Newview.asp" -->
<%
dim t
set t = New ASPTemplate
t.SetTemplatesDir(c_tmp)
t.SetTemplateFile "new_view.html"

t.SetVariableFile "header", "header.html"
t.SetVariableFile "footer", "footer.html"

t.SetVariable "c_tmp",c_tmp
t.SetVariable "c_name",c_name
t.SetVariable "c_key",c_key
t.SetVariable "c_des",c_des
t.SetVariable "productmenu",productmenu
t.SetVariable "footnav",footnav
t.SetVariable "link",link
t.SetVariable "menuname",menuname
t.SetVariable "newsmenu",newsmenu
t.SetVariable "copyright",c_copyright

t.SetVariable "newstitle",newstitle
t.SetVariable "newstime",newstime
t.SetVariable "newsfrom",newsfrom
t.SetVariable "newshit",newshit
t.SetVariable "newsinfo",newsinfo

t.Parse
set t = nothing
%>

