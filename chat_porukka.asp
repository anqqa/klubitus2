<!-- #INCLUDE FILE="dbfuncs.asp" -->
<html>
<head>
	<title>chat . porukka</title>
	<meta http-equiv="refresh" content="30; URL=chat_porukka.asp">
	<style type="text/css">
td {
	font-family: verdana, arial;
	font-size: 11px;
	color: #ffffff;
}
input { 
	background: #328117;
	color: #FFFFFF;
	font-family: Arial;
	font-size: 11px;
}
a { 
	color: #CCFFCC;
	text-decoration: none;
}
a:hover {
	color: #FFFFFF;
	text-decoration: underline
}
	</style>
<%
whois = request.querystring("whois")
if whois <> "" then
	application.lock
	viestit = application("chat_viestit")
	aika = application("chat_aika")
	ip = application("chat_ip")
	nikki = application("chat_nikki")
	selain = application("chat_selain")
	for i = 0 to 48
		viestit(i) = viestit(i + 1)
		aika(i) = aika(i + 1)
		ip(i) = ip(i + 1)
		nikki(i) = nikki(i + 1)
		selain(i) = selain(i + 1)
	next
	aika(49) = now()
	ip(49) = session("ip")
	nikki(49) = session("chat_nikki")
	viestit(49) = "/whois " & whois
	selain(49) = session("selain")
	application("chat_viestit") = viestit
	application("chat_aika") = aika
	application("chat_ip") = ip
	application("chat_nikki") = nikki
	application("chat_selain") = selain
	application.unlock
end if
nimet = session("chat_nimet")
porukkaa = ubound(split(nimet, ",")) + 1
application.lock
application("chat_porukkaa") = porukkaa
application("chatissa") = nimet
application.unlock
nimet = split(nimet, ", ")
for n = 0 to porukkaa - 1
	nimet(n) = "<a href=""chat_porukka.asp?whois=" & nimet(n) & """>" & nimet(n) & "</a>"
next
application.lock
application("chat_porukkaa") = porukkaa
application.unlock
%>
</head>

<body bgcolor="#102908">
<table border="0" cellpadding="0" cellspacing="5">
<tr>
	<td>porukkaa chatissa: <%= porukkaa %> - <%= join(nimet, ", ") %></td>
</tr>
</table>

</body>
</html>