<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = -1000 
if session("nikki") = "" then response.redirect "virhe.asp"

alue = request.querystring("alue")
viesti = request.querystring("viesti")
ketju = request.querystring("ketju")

readyDBCon
if viesti = ketju then
	if session("admin") = "true" then
		Con.Execute "DELETE * FROM messages WHERE id = " & viesti & " OR reply_to = " & viesti
	else
		Con.Execute "DELETE * FROM messages WHERE (id = " & viesti & " OR reply_to = " & viesti & ") AND author_id = " & session("id")
	end if
else
	if session("admin") = "true" then
		Con.Execute "DELETE * FROM messages WHERE id = " & viesti
		Con.Execute "UPDATE messages SET replies = replies - 1 WHERE id = " & ketju
	else
		Con.Execute "DELETE * FROM messages WHERE id = " & viesti & " AND author_id = " & session("id")
		Con.Execute "UPDATE messages SET replies = replies - 1 WHERE id = " & ketju
	end if
end if
Con.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>forum . viesti . poista</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
bueno. viestin poisto onnistu.
<p>
<%
if ketju = viesti then
	response.write "<a href=""forum_alue.asp?alue=" & alue & """ class=""sub1"">takas</a>"
else
	response.write "<a href=""forum_viesti_lue.asp?alue=" & alue & "&viesti=" & ketju & """ class=""sub2"">takas</a>"
end if
%>
	</td>
</tr>
</table>
<div align="right">
<!--WEBBOT bot="HTMLMarkup" startspan ALT="Site Meter" -->
<script type="text/javascript" language="JavaScript">var site="sm3klubitus"</script>
<script type="text/javascript" language="JavaScript1.2" src="http://sm3.sitemeter.com/js/counter.js?site=sm3klubitus">
</script>
<noscript>
<a href="http://sm3.sitemeter.com/stats.asp?site=sm3klubitus" target="_top">
<img src="http://sm3.sitemeter.com/meter.asp?site=sm3klubitus" alt="Site Meter" border=0></a>
</noscript>
<!-- Copyright (c)2000 Site Meter -->
<!--WEBBOT bot="HTMLMarkup" Endspan -->
</div>
</body>
</html>
