<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = -1000 
'if session("nikki") = "" then response.redirect "virhe.asp"

alue = request.form("alue")
viesti = request.form("viesti")
aihe = zeroString(request.form("aihe"))
sisalto = zeroString(trim(request.form("sisalto")))
nimi = zeroString(request.form("nimi"))
nyt = dateadd("h", 1, now())

readyDBCon
Set RS = Server.CreateObject( "ADODB.Recordset" )
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
RS.LockType = adLockOptimistic
RS.Open "SELECT * FROM messages", Con, 2, 2
RS.AddNew
RS("status") = 10
RS("area") = alue
RS("topic") = aihe
if session("nikki") = "" then
	RS("author_id") = -1
	RS("author") = nimi
else
	RS("author_id") = session("id")
	RS("author") = session("nikki")
end if
RS("message") = sisalto
RS("replied") = nyt
RS("last_reply_id") = session("id")
if viesti = "" then RS("reply_to") = -1 else RS("reply_to") = viesti
RS.Update
RS.Close
RS.Open "SELECT * FROM message_areas WHERE id = " & alue, Con, 2, 2
if session("nikki") = "" then
	RS("last_message_author") = nimi
else
	RS("last_message_author") = session("nikki")
end if
RS("last_message_topic") = aihe
RS("last_message_entered") = nyt
RS.Update
RS.Close
if viesti <> "" then
	RS.Open "SELECT * FROM messages WHERE id = " & viesti, Con, 2, 2
	RS("replies") = int(RS("replies")) + 1
	if session("nikki") = "" then
		RS("last_reply_id") = -1
		RS("last_reply") = nimi
	else
		RS("last_reply_id") = session("id")
		RS("last_reply") = session("nikki")
	end if
	RS("replied") = nyt
	RS.Update
	RS.Close
end if
if session("nikki") <> "" then Con.Execute "UPDATE members SET messages = messages+1 WHERE id = " & session("id")
Con.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>forum . viesti . lisäys</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
bueno. viestin lisäys onnistu.
<p>
<%
if viesti = "" then
	response.write "<a href=""forum_alue.asp?alue=" & alue & """ class=""navi"">takas</a>"
else
	response.write "<a href=""forum_viesti_lue.asp?alue=" & alue & "&viesti=" & viesti & "&sivu=69"" class=""navi"">takas</a>"
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
