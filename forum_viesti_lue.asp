<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = -1000 

alue = request.querystring("alue")
viesti = request.querystring("viesti")
sivu = cint(request.querystring("sivu"))
readyDBCon

set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT id, topic, author, entered, message, author_id FROM messages WHERE id = " & viesti & " OR reply_to = " & viesti & " ORDER BY entered"
RS.Open SQL
yht = int(RS.RecordCount)
if yht > 0 then viestit = RS.GetRows()
koko = 20
sivuja = int(yht / koko) + 1
if sivu = 0 then sivu = 1
if sivu > sivuja then sivu = sivuja
RS.Close
SQL = "SELECT id, nick, city, messages FROM members ORDER BY id"
RS.Open SQL
jas = int(RS.RecordCount)
if jas > 0 then rekisteri = RS.GetRows()
RS.Close
SQL = "SELECT id, topic, type FROM message_areas WHERE type < 10 ORDER BY type, id"
RS.Open SQL
alueita = int(RS.RecordCount)
if alueita > 0 then alueet = RS.GetRows()
RS.Close
Con.Close

function etsi(id)
	loyty = -1
	e = 0
	do until loyty > 0 or e = jas
		if int(rekisteri(0, e)) = int(id) then 
			loyty = e
		end if
		e = e + 1
	loop
	etsi = loyty
end function
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>forum . viesti</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<table cellspacing="2" cellpadding="3" border="0">
<tr>
	<td colspan="2">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td align="left">
<% 
for a = 0 to alueita - 1
	if int(alue) = int(alueet(0, a)) then response.write "<a href=""forum_alue.asp?alue=" & alue & """ class=""head"">" & alueet(1, a) & "</a>"
next
%>
	</td>
<form name="alue">
	<td>
<select name="pika" onChange="window.location.href=this.options[this.selectedIndex].value">
<option value="forum.asp">f o r u m i t</option>
<%
for a = 0 to alueita - 1
	if int(alueet(2, a)) = 1 then
		response.write "<option value=""forum_alue.asp?alue=" & alueet(0, a) & """>" & alueet(1, a) & "</option>" & vbNewline
		if int(alue) = int(alueet(0, a)) then
			response.write "<option value=""forum_viesti_lue.asp?alue=" & alue & "&viesti=" & viesti & """ selected> - viesti</option>" & vbNewline
		end if
	end if
next
%>
</select>
	</td>
</form>
	<td align="right">
viestejä yhteensä: <%= yht %> · <a href="forum_viesti_lisaa.asp?alue=<%=alue %>&viesti=<%=viesti %>" class="sub2">vastaa</a>
	</td>
</tr>
</table>
	</td>
</tr>
<%
response.write "<tr><td align=""left"">"
if sivu > 1 then response.write "<a href=""forum_viesti_lue.asp?viesti=" & viesti & "&alue=" & alue & "&sivu=" & (sivu - 1) & """ class=""navi"">edellinen&nbsp;sivu</a>"
response.write "</td><td align=""right"">"
if sivu < sivuja and sivuja > 1 then response.write "<a href=""forum_viesti_lue.asp?viesti=" & viesti & "&alue=" & alue & "&sivu=" & (sivu + 1) & """ class=""navi"">seuraava&nbsp;sivu</a>"
response.write "</td></tr>"
%>
<tr bgcolor="#009900">
	<td class="sub1">kirjoittaja</td>
	<td class="sub1">aihe: <b><% 
response.write viestit(1, 0)
if sivuja > 1 then
	response.write "</b> - sivu: "
	for s = 1 to sivuja
		if sivu = s then 
			response.write "(<b>" & s & "</b>) "
		else
			response.write "<a href=""forum_viesti_lue.asp?alue=" & alue & "&viesti=" & viesti & "&sivu=" & s & """>" & s & "</a> "
		end if
	next
end if
%></td>
</tr>
<%
if yht > 0 then
	alku = (sivu - 1) * koko
	loppu = sivu * koko - 1
	if loppu > (yht - 1) then loppu = yht - 1
	for i = alku to loppu
		teksti = replace(viestit(4, i), vbNewline, "<br>")
		teksti = replace(teksti, ":x-", "<img align=absmiddle src=smiley/AR15firing.gif>")
		teksti = replace(teksti, ":X-", "<img align=absmiddle src=smiley/AR15firing.gif>")
		teksti = replace(teksti, "B)", "<img align=absmiddle src=smiley/cool.gif>")
		teksti = replace(teksti, ";)D", "<img align=absmiddle src=smiley/deal.gif>")
		teksti = replace(teksti, ":o.", "<img align=absmiddle src=smiley/drooling3.gif>")
		teksti = replace(teksti, "8(", "<img align=absmiddle src=smiley/puppy_dog_eyes.gif>")
		teksti = replace(teksti, ";(", "<img align=absmiddle src=smiley/sad2.gif>")
		teksti = replace(teksti, ":D", "<img align=absmiddle src=smiley/sweat.gif>")
		teksti = replace(teksti, "=D", "<img align=absmiddle src=smiley/lol.gif>")
		teksti = replace(teksti, "d:)", "<img align=absmiddle src=smiley/ymca.gif>")
		teksti = replace(teksti, "q:)", "<img align=absmiddle src=smiley/ymca.gif>")
		teksti = replace(teksti, ":o)", "<img align=absmiddle src=smiley/clown.gif>")
		teksti = replace(teksti, ":/", "<img align=absmiddle src=smiley/indiff.gif>")
		teksti = replace(teksti, ":I", "<img align=absmiddle src=smiley/indiff.gif>")
		teksti = replace(teksti, ":666", "<img align=absmiddle src=smiley/devil.gif>")
		teksti = replace(teksti, "X(", "<img align=absmiddle src=smiley/dead.gif>")
		teksti = replace(teksti, "x(", "<img align=absmiddle src=smiley/dead.gif>")
		teksti = replace(teksti, ":B", "<img align=absmiddle src=smiley/bucked.gif>")
		teksti = replace(teksti, ":s", "<img align=absmiddle src=smiley/badteeth.gif>")
		teksti = replace(teksti, ":S", "<img align=absmiddle src=smiley/badteeth.gif>")
		teksti = replace(teksti, ":x)", "<img align=absmiddle src=smiley/nibble6.gif>")
		teksti = replace(teksti, "x)", "<img align=absmiddle src=smiley/loveeyes.gif>")
		teksti = replace(teksti, "X)", "<img align=absmiddle src=smiley/loveeyes.gif>")
		teksti = replace(teksti, "o:)", "<img align=absmiddle src=smiley/mpangel.gif>")
		teksti = replace(teksti, "O:)", "<img align=absmiddle src=smiley/mpangel.gif>")
		teksti = replace(teksti, ":oo", "<img align=absmiddle src=smiley/fruit.gif>")
		teksti = replace(teksti, ":Q", "<img align=absmiddle src=smiley/rasta.gif>")
		teksti = replace(teksti, ":)B", "<img align=absmiddle src=smiley/moon.gif>")
		teksti = replace(teksti, "x:(", "<img align=absmiddle src=smiley/newburn.gif>")
		teksti = replace(teksti, ":.(", "<img align=absmiddle src=smiley/mecry.gif>")
		teksti = replace(teksti, ":(", "<img align=absmiddle src=smiley/sad.gif>")
		teksti = replace(teksti, ":x", "<img align=absmiddle src=smiley/smiley_kiss.gif>")
		teksti = replace(teksti, ":)", "<img align=absmiddle src=smiley/smile.gif>")
		teksti = replace(teksti, "s:o", "<img align=absmiddle src=smiley/gorgeous.gif>")
		teksti = replace(teksti, ":pb", "<img align=absmiddle src=smiley/readpb.gif>")
		teksti = replace(teksti, ":o", "<img align=absmiddle src=smiley/oogle.gif>")
		teksti = replace(teksti, ":O", "<img align=absmiddle src=smiley/oogle.gif>")
		teksti = replace(teksti, ":P", "<img align=absmiddle src=smiley/buck.gif>")
		teksti = replace(teksti, ":p", "<img align=absmiddle src=smiley/buck.gif>")
		teksti = replace(teksti, ";P", "<img align=absmiddle src=smiley/hehehmn.gif>")
		teksti = replace(teksti, ";p", "<img align=absmiddle src=smiley/hehehmn.gif>")
		teksti = replace(teksti, ";D", "<img align=absmiddle src=smiley/evil_laughterpurple.gif>")
		teksti = replace(teksti, "8o", "<img align=absmiddle src=smiley/shocked.gif>")
		teksti = replace(teksti, "=)", "<img align=absmiddle src=smiley/tongue.gif>")
		teksti = replace(teksti, ":9", "<img align=absmiddle src=smiley/tongue.gif>")
		teksti = replace(teksti, ";)", "<img align=absmiddle src=smiley/wink2.gif>")
		teksti = replace(teksti, ":x", "<img align=absmiddle src=smiley/wink.gif>")
		teksti = replace(teksti, "3-o", "<img align=absmiddle src=smiley/sex.gif>")
		teksti = replace(teksti, "3-O", "<img align=absmiddle src=smiley/sex.gif>")
		teksti = replace(teksti, "=/", "<img align=absmiddle src=smiley/smiley_blush.gif>")
		response.write "<tr>" & vbNewline
		response.write "<td class=""sub2"" valign=""top"">" & vbNewline
		m = etsi(viestit(5, i))
		if m < 0 then
			response.write "<b>" & viestit(2, i) & "</b><br>" & vbNewline
			response.write "<i>vierailija</i>" & vbNewline
		else
			response.write "<b>" & rekisteri(1, m) & "</b>&nbsp;(#" & rekisteri(0, m) & ")<br>" & vbNewline
			response.write "viestejä: " & rekisteri(3, m) & "<br>" & vbNewline
			response.write "kaupunki:&nbsp;" & rekisteri(2, m) & "" & vbnewline
		end if
		if session("admin") = "true" or viestit(2, i) = session("nikki") then
			response.write "<p><a href=""forum_viesti_poista.asp?ketju=" & viestit(0, 0) & "&viesti=" & viestit(0, i) & "&alue=" & alue & """>poista</a>"
		end if
		response.write "</td>" & vbnewline
		response.write "<td valign=""top""><div class=""sub1"">" & viestit(3, i) & " (#" & (i + 1) & ")</div><br>" & vbnewline & teksti & "</td>" & vbNewline
		response.write "</tr>" & vbNewline
		response.write "<tr>" & vbNewline
		response.write "<td bgcolor=""#009900"" height=""1"" colspan=""2""></td>" & vbNewline
		response.write "</tr>" & vbNewline
	next
	response.write "<tr><td align=""left"">"
	if sivu > 1 then response.write "<a href=""forum_viesti_lue.asp?viesti=" & viesti & "&alue=" & alue & "&sivu=" & (sivu - 1) & """ class=""navi"">edellinen sivu</a>"
	response.write "</td><td align=""right"">"
	if sivu < sivuja and sivuja > 1 then response.write "<a href=""forum_viesti_lue.asp?viesti=" & viesti & "&alue=" & alue & "&sivu=" & (sivu + 1) & """ class=""navi"">seuraava sivu</a>"
	response.write "</td></tr>"
end if
%>
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
