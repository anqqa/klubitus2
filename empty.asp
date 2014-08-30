<%
if request.servervariables("HTTP_REFERER") = "http://www.anqqa.com/klubitus/" then
%>
<html><head><title>empty</title></head>
<frameset cols="200,*" border="0" frameborder="0" framespacing="0"><frame name="" src="empty.html" marginwidth="0" marginheight="0" scrolling="no" frameborder="0" noresize><frame name="" src="sana2.asp?koe=testi" marginwidth="0" marginheight="0" scrolling="no" frameborder="0" noresize></frameset>
</html>
<%
else
	response.redirect "empty.html"
end if
%>
