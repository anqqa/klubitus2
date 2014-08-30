<!-- #INCLUDE FILE="dbfuncs.asp" -->
<html>
<head>
	<title>chat</title>
</head>

<frameset rows="*,40,20" frameborder="0" border="0">
	    <frame name="viestit" src="chat_viestit.asp" marginwidth="0" marginheight="0" scrolling="auto" frameborder="0" noresize>
<%
if session("admin") = "true" then 
	response.write "<frameset cols=""90%,*"">"
else
	response.write "<frameset cols=""100%,*"">"
end if
%>
        <frame name="kirjoita" src="chat_kirjoita.asp" marginwidth="0" marginheight="0" scrolling="no" frameborder="0" noresize>
        <frame name="poll" src="chat_poll.asp" marginwidth="0" marginheight="0" scrolling="no" frameborder="0" noresize>
    </frameset>
    <frame name="porukka" src="chat_porukka.asp" marginwidth="0" marginheight="0" scrolling="no" frameborder="0" noresize>
</frameset>

</html>