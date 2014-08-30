<% session("index") = true %>
<!-- #INCLUDE FILE="dbfuncs.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>K L U B I T U S</title>
	<link rel="shortcut icon" href="http://www.anqqa.com/klubitus/favicon.ico">
	<meta name="description" content="Klubituksesta löydät kaiken mitä tarvitset tietää Helsingin bileistä: klubit, mestat, kaupat jne.">
	<meta name="keywords" content="klubitus,helsinki,klubi,klubit,bileet,bailut,bile,chat,kuvat,keskustelu,forum,radio,uutiset,ravintola,viima,kerma,hype,agenda,labyrinth,vertigo,dj,nosturi,soda,opas,klubiopas,ever,life,pussy,pump">
<!--	<meta http-equiv="Page-Exit" content="blendTrans(Duration=2.0)">-->
<%
session("kaupunki") = request.querystring("kaupunki")
if request.querystring("logout") = "true" then
	session("nikki") = ""
	session("id") = ""
	session("email") = ""
	session("admin") = ""
else
	if request.cookies("klubitus")("id") <> "" then
		session("id") = request.cookies("klubitus")("id")
		session("nikki") = request.cookies("klubitus")("nikki")
		session("email") = request.cookies("klubitus")("email")
	end if
end if

if (session("nikki") <> "") then
	readyDBCon
	Con.Execute "UPDATE members SET visits = visits+1 WHERE id = " & session("id")
	Con.Execute "UPDATE members SET lastvisit = Now() WHERE id = " & session("id")
	set RS = Con.Execute("SELECT config FROM members WHERE id = " & session("id"))
	configgi = RS("config")
	if configgi <> "" then
		configgi = replace(configgi, "kaupungit=+", "")
'		session("config") = replace(configgi, "nayta=+", "")
	end if
	Con.Close
end if
%>
</head>

<!-- frames -->
<frameset rows="90,*,90" frameborder=0 border=0>
    <frame name="header" src="navi.asp" marginwidth="0" marginheight="0" scrolling="no" frameborder="no" noresize>
	<frameset cols="150,*" frameborder=0 border=0>
	    <frame name="menu" src="login.asp" marginwidth="5" marginheight="10" scrolling="auto" frameborder="no" noresize>
<% if request.querystring("sivu") = "" then %>
	    <frame name="main" src="bileet.asp" marginwidth="10" marginheight="10" scrolling="auto" frameborder="no" noresize>
<% else %>
	    <frame name="main" src="<%= request.querystring("sivu") %>" marginwidth="10" marginheight="10" scrolling="auto" frameborder="no" noresize>
<% end if %>
	</frameset>
	<frameset cols="150,*" frameborder=0 border=0>
	    <frame name="empty" src="empty.html" marginwidth="10" marginheight="10" scrolling="no" frameborder="no" noresize>
	    <frame name="footer" src="sana.asp" marginwidth="5" marginheight="5" scrolling="no" frameborder="no" noresize>
	</frameset>
</frameset>
