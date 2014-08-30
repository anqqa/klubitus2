<!-- #INCLUDE FILE="dbfuncs.asp" -->
<% 
   Response.Expires = 0
   Response.ExpiresAbsolute = Now() - 1
   Response.AddHeader "cache-control", "private"
   Response.AddHeader "pragma", "no-cache"
   Response.CacheControl="no-cache"
%>
<html>
<head>
	<title>rekkaus</title>
<link rel="stylesheet" href="keltainen.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_keltainen.jpg" bgproperties="fixed">

<%
nikki = request.form("nikki")
salasana = request.form("salasana")
salasana2 = request.form("salasana2")
kuvaus = zeroString(request.form("kuvaus"))
signature = zeroString(request.form("signature"))
nimi = request.form("nimi")
email = request.form("email")
kotisivu = formatURL(request.form("kotisivu"))
kaupunki = zeroString(request.form("kaupunki"))
synt_paiva = request.form("synt_paiva")
synt_kuukausi = request.form("synt_kuukausi")
synt_vuosi = request.form("synt_vuosi")

' salasana unohtunu?
if (nikki <> "") and (salasana = "") and (nimi = "") and (email = "") then
	readyDBCon
	SQL = "SELECT password, name, email FROM members WHERE nick ='" & nikki & "'"
	set RS = Con.Execute(SQL)
	if RS.BOF and RS.EOF then
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td><b>erhe!</b><br>etpä taida olla rekisteröityny, ei ainakaan löytyny moisella käyttäjätunnuksella ketään..</td>
</tr>
</table>
<%
	else
		mail "Klubitus", "klubitus@anqqa.com", nikki & "(" & RS("name") & ")", RS("email"), "salasana", RS("password")
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>salasana meilattu..</td>
</tr>
</table>
<%
	end if
	RS.Close
	Con.Close
else
' puuttuuko kohtia?
	dim puuttuu
	if nikki = "" then puuttuu = "käyttäjätunnus "
	if (salasana = "") or (salasana2 = "") then puuttuu = puuttuu + "salasana "
	if nimi = "" then puuttuu = puuttuu + "nimi "
	if email = "" or not emailOK(email) then puuttuu = puuttuu + "email "
	if puuttuu <> "" then
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
<b>erhe!</b><br>et täyttäny kaikkia vaadittuja kohtia: <b><%=puuttuu %></b>
<p>
<a href="javascript:history.go(-1)" class="navi">takas</a>
	</td>
</tr>
</table>
<%
' salasana kunnossa?
	elseif salasana <> salasana2 then
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
<b>erhe!</b><br>salasanat ei täsmää..
<p>
<a href="javascript:history.go(-1)" class="navi">takas</a>

	</td>
</tr>
</table>
<%
' kaikki okay
	else
' tietojen lisäys, samojen tarkistus
		readyDBCon
		SQL = "SELECT id FROM members WHERE nick ='" & nikki & "' OR email = '" & email & "'"
		set RS = Con.Execute(SQL)
		if not (RS.BOF and RS.EOF) then
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
<b>erhe!</b><br>jollain on sama käyttäjätunnus tai e-mail..
<p>
<a href="javascript:history.go(-1)" class="navi">takas</a>
	</td>
</tr>
</table>
<%
			RS.Close
			Con.Close
		else	
' rekisteröinti!
			readyDBCon
			Set RS = Server.CreateObject( "ADODB.Recordset" )
			RS.ActiveConnection = Con
			RS.CursorType = adOpenStatic
			RS.LockType = adLockOptimistic
			RS.Open "SELECT * FROM members", Con, 2, 2
			RS.AddNew
			RS("name") = nimi
			RS("email") = email
			RS("homepage") = kotisivu
			RS("desc") = kuvaus
			RS("dob_day") = synt_paiva
			RS("dob_month") = synt_kuukausi
			RS("dob_year") = synt_vuosi
			RS("city") = kaupunki
			RS("nick") = nikki
			RS("password") = salasana
			RS("signature") = signature
			RS("config") = ""
			RS.Update
			RS.Close
			Con.Close
'			mail "Klubitus", "klubitus@anqqa.com", "rekkaus", "klubitus@anqqa.com", "rekkaus", nikki
%>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td>
bueno. rekisteröityminen luonnistu, nyt pystyt esim
<ul>
<li>ottamaan osaa keskusteluun
<li>lisäämään ja muokkaamaan pileitä
<li>säätämään bilelistan muotoa
<li>muuttamaan tietoja <b>asetukset</b> valikosta
</ul>
<b>huom!</b><br>
jos unohdat salasanan, rekisteröidy täyttämällä <u>pelkkä</u> käyttäjänimi ja saat salasanan meilissä..
</td>
</tr>
</table>
<%
		end if
	end if
end if
%>
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
