<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.CacheControl="no-cache"

if session("admin") <> "true" then response.redirect "virhe.asp"
viikko = datepart("ww", now())
alkaa = datevalue(dateadd("d", 1-datepart("w", now(), vbmonday), now()))
loppuu = datevalue(dateadd("d", 7-datepart("w", now(), vbmonday), now()))
uutiset = request.form("uutiset")
header = "klubitus news, viikko " & viikko & " (" & paiva2(alkaa) & " - " & paiva2(loppuu) & ")"
%>

<html>
<head>
	<title>news letter</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<%
readyDBCon
set RS = Server.CreateObject( "ADODB.Recordset" )
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT name, nick, email, city, join, lastvisit, visits FROM members WHERE (type > 0) AND (config LIKE '%newsletter%')"
RS.Open SQL
yht = int(RS.RecordCount)
rekisteri = RS.GetRows() ' kenttä, numero
RS.Close
sivunkoko = 5
sivuja =  int(yht / sivunkoko) + 1
if yht > 0 then
	response.write "<table cellspacing=""5"" cellpadding=""0"" width=""100%"" border=""0"">"
	response.write "<tr><td colspan=""4"">yhteensä: " & yht & "</td></tr>"
	for i = 0 to yht - 1 
%>
<tr>
	<td><%= rekisteri(0, i) %> (<%= rekisteri(1, i) %>)</td>
	<td><%= rekisteri(3, i) %></td>
	<td>rekisteröitynyt <%= rekisteri(4, i) %></td>
	<td>viimeksi <%= rekisteri(5, i) %> (yht. <%= rekisteri(6, i) %>)</td>
</tr>
<% 
	next 
	if uutiset <> "" then
		for s = 1 to sivuja
			Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
			Mailer.FromName = "klubitus news"
			Mailer.FromAddress = "news@klubitus.com"
'			Mailer.RemoteHost = "mail.inet.fi"
			Mailer.RemoteHost = "smtpgw.activeisp.com"
			Mailer.Subject = header
			Mailer.Organization = "klubitus"
'			Mailer.AddRecipient "klubitus news", "news@klubitus.com"
			Mailer.CharSet = 2
			response.write "<tr><td colspan=4>"
			Mailer.BodyText = uutiset
			alku = sivunkoko * (s - 1)
			loppu = sivunkoko * s
			if loppu > (yht - 1) then loppu = yht - 1
			for i = alku to loppu
				Mailer.AddBCC rekisteri(1, i), rekisteri(2, i)
			next
			response.write "<br><br>" & s & ". " & (alku + 1) & "-" & loppu & ": "
			if Mailer.SendMail then
				Response.Write "uutislehtinen lähetetty.."
			else
				Response.Write "ei lähteny. virhe: " & Mailer.Response
			end if
			response.write "</td></tr>"
			set Mailer = nothing
		next
		response.write "<tr><td colspan=""4"">" & replace(uutiset, vbnewline, "<br>") & "</td></tr>" & vbnewline
	end if
	response.write "</table>"
end if 

if uutiset = "" then
	SQL = "SELECT day, name, place, city, dj, hours_from, hours_to, age, price, music, info, day, month, year FROM parties ORDER BY year, month, day, city, name"
	RS.Open SQL
	yht = int(RS.RecordCount)
	bileet = RS.GetRows() ' kenttä, numero
	RS.Close
	Con.Close
	yht2 = 0
	tanaan = -1
	newsletter = ""
	for i = 0 to yht - 1
		pvm = datevalue(bileet(11, i) & "/" & bileet(12, i) & "/" & bileet(13, i))
		if datepart("ww", pvm, vbmonday) = viikko then
			if datepart("w", pvm, vbmonday) <> tanaan then
				tanaan = datepart("w", pvm, vbmonday)
				newsletter = newsletter & vbnewline & paivat(datepart("w", pvm, vbmonday) - 1) & " " & formatdatetime(pvm, 2) & vbnewline
				newsletter = newsletter & string(50, "-") & vbnewline
			end if
			newsletter = newsletter & bileet(1, i) & " @ " & bileet(2, i) & ", " & bileet(3, i) & " ("
			if bileet(5, i) > -1 then 
				newsletter = newsletter & zero(bileet(5, i)) & "-"
				if bileet(6, i) > -1 then newsletter = newsletter & zero(bileet(6, i)) & " " else newsletter = newsletter & "> "
			end if
			if bileet(7, i) > 0 then newsletter = newsletter & "k" & bileet(7, i) & " "
			if bileet(8, i) > -1 then newsletter = newsletter & bileet(8, i) & "mk"
			newsletter = rtrim(newsletter) & ")" & vbnewline
			if len(bileet(4, i)) > 1 then newsletter = newsletter & "dj: " & bileet(4, i) & vbnewline
			if len(bileet(10, i)) > 1 then newsletter = newsletter & "lisätiedot: " & replace(bileet(10, i), vbnewline, ". ") & vbnewline
			if len(bileet(9, i)) > 1 then newsletter = newsletter & bileet(9, i) & vbnewline
			newsletter = newsletter & vbnewline
			yht2 = yht2 + 1
		end if
	next
%>
<br>
<form action="newsletter.asp" method="post">
<table cellpadding="0" cellspacing="5" border="0" width="100%">
<tr>
	<td><textarea cols="80" rows="100" name="uutiset">
<%
response.write string(50, "+") & vbnewline
response.write header & vbnewline
response.write string(50, "+") & vbnewline & vbnewline
response.write "· uutinen 1" & vbnewline
response.write string(50, "-") & vbnewline & vbnewline & vbnewline & vbnewline
response.write "· uutinen 2" & vbnewline
response.write string(50, "-") & vbnewline & vbnewline & vbnewline & vbnewline
response.write string(50, "+") & vbnewline
response.write "bileet" & vbnewline
response.write string(50, "+") & vbnewline & vbnewline
response.write newsletter
%>
	</textarea>
	</td>
</tr>
<tr>
	<td><input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok"></td>
</tr>
</table>
</form>
<% end if %>
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