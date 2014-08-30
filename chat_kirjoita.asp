<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
if request.querystring("uusi") = "true" then session("chat_nikki") = ""
if session("nikki") <> "" then session("chat_nikki") = session("nikki")
session("selain") = request.servervariables("HTTP_USER_AGENT")
session("ip") = request.servervariables("REMOTE_ADDR")
if session("chat_nikki") <> "" then

viesti = request.form("viesti")
application.lock
dim viestit, aika, ip, nikki
if application("chat_ok") <> "true" then
	redim viestit(50), selain(50), aika(50), ip(50), nikki(50), porukka(50), porukka_ip(50), porukka_aika(50), porukka_selain(50), alias(20), aliasto(20), block(20)
	application("chat_viestit") = viestit
	application("chat_aika") = aika
	application("chat_ip") = ip
	application("chat_nikki") = nikki
	application("chat_selain") = selain
	application("chat_porukka") = porukka
	application("chat_porukka_ip") = porukka_ip
	application("chat_porukka_aika") = porukka_aika
	application("chat_porukka_selain") = porukka_selain
	application("chat_alias") = alias
	application("chat_aliasto") = aliasto
	application("chat_block") = block
	application("chat_ok") = "true"
else
	viestit = application("chat_viestit")
	aika = application("chat_aika")
	ip = application("chat_ip")
	nikki = application("chat_nikki")
	selain = application("chat_selain")
	block = application("chat_block")
end if
if viesti = "" then 
'	if session("nikki") = "" then session("chat_nikki") = "vierailija" else session("chat_nikki") = session("nikki")
	if session("nikki") = "" then viesti = "/me (vieras) tuli.." else viesti = "/me tuli.."
	session("riveja") = 10
	session("uusin") = dateadd("yyyy", -1, now())
end if
'komento: /last x
if lcase(left(viesti, 5)) = "/last" then 
	if len(viesti) > 5 then riveja = int(right(viesti, len(viesti) - 6)) else riveja = 10
	session("uusin") = dateadd("yyyy", -1, now())
	if riveja > 49 then riveja = 49
	session("riveja") = riveja
else
	blocked = false
	for i = 0 to 19
		if block(i) = session("chat_nikki") then blocked = true
	next
	if not blocked then
		for i = 0 to 48
			viestit(i) = viestit(i + 1)
			aika(i) = aika(i + 1)
			ip(i) = ip(i + 1)
			nikki(i) = nikki(i + 1)
			selain(i) = selain(i + 1)
		next
		viesti = replace(viesti, """", "&quot;")
		viesti = replace(viesti, "'", "`")
		viesti = replace(viesti, "<", "&laquo;")
		viesti = replace(viesti, ">", "&raquo;")
		viesti = replace(viesti, "&lt;", "&laquo;")
		viesti = replace(viesti, "&lt", "&laquo;")
		viesti = replace(viesti, "&LT;", "&laquo;")
		viesti = replace(viesti, "&LT", "&laquo;")
		viesti = replace(viesti, "&gt;", "&raquo;")
		viesti = replace(viesti, "&gt", "&raquo;")
		viesti = replace(viesti, "&GT;", "&raquo;")
		viesti = replace(viesti, "&GT", "&raquo;")
		aika(49) = now()
		ip(49) = session("ip")
		nikki(49) = session("chat_nikki")
		viestit(49) = viesti
		selain(49) = session("selain")
	end if
end if
application("chat_viestit") = viestit
application("chat_aika") = aika
application("chat_ip") = ip
application("chat_nikki") = nikki
application("chat_selain") = selain
application.unlock
else
	teksti = "valitse nimi: "
	nimi = request.form("viesti")
	if nimi <> "" then
		nimi = replace(nimi, "8o", "8O")
		nimi = replace(nimi, "3-o", "3-0")
		nimi = replace(nimi, ":", "")
		nimi = replace(nimi, ";", "")
		nimi = replace(nimi, "=", "")
		nimi = replace(nimi, " ", "_")
		nimi = replace(nimi, "<", "")
		nimi = replace(nimi, ">", "")
		nimi = replace(nimi, ")", "")
		nimi = replace(nimi, "(", "")
		nimi = replace(nimi, "o/", "0/")
		nimi = replace(nimi, """", "")
		nimi = replace(nimi, "'", "")
		nimi = replace(nimi, "anqqa", "angga")
		readyDBCon
		SQL = "SELECT id FROM members WHERE nick ='" & nimi & "'"
		set RS = Con.Execute(SQL)
		if not (RS.BOF and RS.EOF) then
			teksti = "rekisteröity nimi.<br>&nbsp;valitse joku muu: "
		else
			session("chat_nikki") = nimi
			response.redirect "chat_kirjoita.asp"
		end if
		RS.Close
		Con.Close
	end if
end if
%>
<html>
<head>
	<title>chat . viestit</title>
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
	</style>
	<script language="javascript">
<!--
function check() {
	if (document.kirjoita.viesti.value != '') document.kirjoita.submit();
}
//-->
	</script>
</head>

<body bgcolor="#102908" onload="document.all.viesti.focus()">
<table height="100%" width="100%" cellpadding="0" cellspacing="0" border="0">
<form name="kirjoita" action="chat_kirjoita.asp" method="post" onsubmit="check(); return false">
<tr>
	<td valign="middle" width="370">&nbsp;<% if session("chat_nikki") = "" then response.write teksti else response.write "viesti: " %><input type="text" name="viesti" size="50"></td>
	<td width="50">&nbsp;&nbsp;<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok"></td>
	<% if session("chat_nikki") <> "" then %><td>&nbsp;/help = <i>apua</i><br>&nbsp;/last x = <i>edelliset x viestiä</i></td><% end if %>
	<td align="right">
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
	</td>
</tr>
</form>
</table>
</body>
</html>