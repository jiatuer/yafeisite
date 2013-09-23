<!-- #include file="Include/conn.asp" -->
<%
function g(t)
	if len(t)>19 then
		g=left(t,19)&"бн"
	else
		g=t
	end if
end function

pic=""
link=""
imgtext=""

	sql="select top 5 * from xin_Index where img<>'' order by id desc"
	Set rs=conn.execute(sql)
	do while not rs.eof
		pic=pic&"|"&rs("img")
		link=link&"|"&"new_view.asp?id="&rs("id")
		imgtext=imgtext&"|"&g(rs("titles"))
		rs.MoveNext
	loop
	rs.Close
	set rs=nothing


pic=right(pic,len(pic)-1)
link=right(link,len(link)-1)
imgtext=right(imgtext,len(imgtext)-1)
%>
<a target=_blank href="javascript:goUrl()"> 
<span class="f14b">
<script type="text/javascript">

 var focus_width=445
 var focus_height=220
 var text_height=0
 var swf_height = focus_height+text_height
 
 var pics="<%=pic%>"
 var links="<%=link%>"
 var texts="<%=imgtext%>"
 
document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ focus_width +'" height="'+ swf_height +'">');
document.write('<param name="allowScriptAccess" value="sameDomain"><param name="movie" value="images/playswf.swf"><param name=wmode value=transparent><param name="quality" value="high">');
document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
document.write('<param name="FlashVars" value="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'">');
document.write('<embed src="images/playswf.swf" wmode="opaque" FlashVars="pics='+pics+'&links='+links+'&texts='+texts+'&borderwidth='+focus_width+'&borderheight='+focus_height+'&textheight='+text_height+'" menu="false" bgcolor="#DADADA" quality="high" width="'+ focus_width +'" height="'+ swf_height +'" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />');  document.write('</object>');
//-->
</script> 
</span></a><span id=focustext class=f14b> </span>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style><body>
</body>
</html>
