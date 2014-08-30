<!-- #INCLUDE FILE="dbfuncs.asp" -->
<% 
Response.Expires = 0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control", "private"
Response.AddHeader "pragma", "no-cache"
Response.CacheControl="no-cache"
Response.Buffer = true

salasana = request.form("salasana")
salasana2 = request.form("salasana2")
kuvaus = zeroString(request.form("kuvaus"))
signature = zeroString(request.form("signature"))
email = request.form("email")
kotisivu = formatURL(request.form("kotisivu"))
kaupunki = zeroString(request.form("kaupunki"))
synt_paiva = request.form("synt_paiva")
synt_kuukausi = request.form("synt_kuukausi")
synt_vuosi = request.form("synt_vuosi")

tallenna = request.form("tallenna")
kaupungit = request.form("kaupungit")
nayta = request.form("nayta")
newsletter = request.form("newsletter")

muutos = request.form("muutos")
%>

<html>
<head>
	<title>asetukset</title>
	<link rel="stylesheet" href="punainen.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_punainen.jpg" bgproperties="fixed">

<%
readyDBCon
feelua = false
if muutos = "true" then
	if salasana <> "" then
		if salasana <> salasana2 then
			feelua = true
%>
<table cellspacing="0" cellpadding="5" border="0">
<tr>
	<td><b>erhe!</b><br>salasanat ei t‰sm‰‰..</td>
</tr>
</table>
<%
		else
			salasana3 = salasana
		end if
	end if
	if not emailOK(email) then
		feelua = true
%>
<table cellspacing="0" cellpadding="5" border="0">
<tr>
	<td><b>erhe!</b><br>e-mail ei ok..</td>
</tr>
</table>
<%
	end if
	if not feelua then
		SQL = "SELECT password, name, email, homepage, desc, dob_day, dob_month, dob_year, city, nick, signature, config, id FROM members WHERE nick = '" & session("nikki") & "'"
		Set RS = Server.CreateObject( "ADODB.Recordset" )
		RS.ActiveConnection = Con
		RS.CursorType = adOpenStatic
		RS.LockType = adLockOptimistic
		RS.Open SQL, Con, 3, 3
		RS("homepage") = kotisivu
		RS("desc") = kuvaus
		RS("dob_day") = synt_paiva
		RS("dob_month") = synt_kuukausi
		RS("dob_year") = synt_vuosi
		RS("city") = kaupunki
		RS("signature") = signature
		configgi = "kaupungit=" & kaupungit & "+nayta=" & nayta & "+" & newsletter
		configgi = replace(configgi, "kaupungit=+", "")
		configgi = replace(configgi, "nayta=+", "")
		if configgi <> "" then RS("config") = configgi
		session("config") = RS("config")
		if salasana3 <> "" then RS("password") = salasana3
		if emailOK(email) then RS("email") = email
		if tallenna = "true" then
			response.cookies("klubitus")("id") = int(RS("id"))
			response.cookies("klubitus")("nikki") = session("nikki")
			response.cookies("klubitus")("email") = RS("email")
			response.cookies("klubitus").expires = date() + 365
		else
			response.cookies("klubitus")("id") = ""
			response.cookies("klubitus").expires = date() - 1
		end if
		RS.Update
		RS.Close
%>
<table cellspacing="0" cellpadding="5" border="0">
<tr>
	<td>jees, tiedot muutettu..</td>
</tr>
</table>
<%
	end if
end if

SQL = "SELECT name, email, homepage, desc, dob_day, dob_month, dob_year, city, nick, signature, config FROM members WHERE nick = '" & session("nikki") & "'"
set RS = Con.Execute(SQL)
config = RS("config")
%>
<form action="asetukset.asp" method="post">
<table cellspacing="0" cellpadding="0" border="0">
<tr>
	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td class="selite"><b><i>k‰ytt‰j‰tunnus:</i></b></td>
		<td><%=RS("nick") %></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><b>salasana:</b></td>
		<td><input type="password" maxlength="20" size="30" value="" name="salasana"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><b>uudestaan:</b></td>
		<td><input type="password" maxlength="20" size="30" value="" name="salasana2"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite" valign="top"><i>kuvaus/<br>vapaa sana:</i></td>
		<td><textarea name="kuvaus" rows="5" cols="30"><%=RS("desc") %></textarea></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite" valign="top"><i>signature:</i></td>
		<td><textarea name="signature" cols="30"><%=RS("signature") %></textarea></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><b>nimi:</b></td>
		<td><%=RS("name") %></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><b>e-mail:</b></td>
		<td><input type="text" maxlength="60" size="30" name="email" value="<%=RS("email") %>"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><i>kotisivu:</i></td>
		<td><input type="text" maxlength="60" size="30" name="kotisivu" value="<%=RS("homepage") %>"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><i>kaupunki:</i></td>
		<td><input type="text" maxlength="30" size="30" name="kaupunki" value="<%=RS("city") %>"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">syntym‰aika:</td>
		<td>
		<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td>
<select name="synt_paiva">
<option value="0" <% if RS("dob_day") = 0 then response.write("selected") %>>--</option>
<%
for i = 1 to 31
	if RS("dob_day") = i then
		response.write "<option value='" & i & "' selected>" & i & ".</option>"
	else
		response.write "<option value='" & i & "'>" & i & ".</option>"
	end if
next
%>
</select>
			</td>
			<td>
<select name="synt_kuukausi">
<option value="0" <% if RS("dob_month") = 0 then response.write("selected") %>>--</option>
<option value="1" <% if RS("dob_month") = 1 then response.write("selected") %>>tammikuuta</option>
<option value="2" <% if RS("dob_month") = 2 then response.write("selected") %>>helmikuuta</option>
<option value="3" <% if RS("dob_month") = 3 then response.write("selected") %>>maaliskuuta</option>
<option value="4" <% if RS("dob_month") = 4 then response.write("selected") %>>huhtikuuta</option>
<option value="5" <% if RS("dob_month") = 5 then response.write("selected") %>>toukokuuta</option>
<option value="6" <% if RS("dob_month") = 6 then response.write("selected") %>>kes‰kuuta</option>
<option value="7" <% if RS("dob_month") = 7 then response.write("selected") %>>hein‰kuuta</option>
<option value="8" <% if RS("dob_month") = 8 then response.write("selected") %>>elokuuta</option>
<option value="9" <% if RS("dob_month") = 9 then response.write("selected") %>>syyskuuta</option>
<option value="10" <% if RS("dob_month") = 10 then response.write("selected") %>>lokakuuta</option>
<option value="11" <% if RS("dob_month") = 11 then response.write("selected") %>>marraskuuta</option>
<option value="12" <% if RS("dob_month") = 12 then response.write("selected") %>>joulukuuta</option>
</select>
			</td>
			<td>
<select name="synt_vuosi">
<option value="0" <% if RS("dob_year") = 0 then response.write("selected") %>>--</option>
<% for i = 1990 to 1955 step -1 
	if RS("dob_year") = i then
		response.write "<option value='" & i & "' selected>" & i & "</option>"
	else
		response.write "<option value='" & i & "'>" & i & "</option>"
	end if
next %>
</select>
			</td>
		</tr>
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	</table>
	</td>

	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td colspan="2">
<input type="checkbox" name="newsletter" value="newsletter" <% if instr(config, "newsletter") > 0 then response.write "checked" %>> newsletter<br>
jos haluat maanantaisin s‰hkˆpostilla tiedot kaikista viikon bileist‰ (under development)
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2">
<input type="checkbox" name="tallenna" value="true" <% if request.cookies("klubitus")("id") <> "" then response.write "checked" %>> automaattinen login<br>
(huom! eli siis tallentaa nimen ja salasanan cookiella t‰h‰n selaimeen, jolloin automaattisesti kirjaudut sis‰‰n t‰nne tullessas, <u>‰l‰</u> siis k‰yt‰ jos muilla p‰‰sy koneelle tms)
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td>
n‰yt‰ kaupungit:<br>
<input type="checkbox" name="kaupungit" value="helsinki"<% if instr(config, "helsinki") > 0 or RS("config") = "" then response.write "checked" %>> helsinki<br>
<input type="checkbox" name="kaupungit" value="j‰rvenp‰‰"<% if instr(config, "j‰rvenp‰‰") > 0 or RS("config") = "" then response.write "checked" %>> j‰rvenp‰‰<br>
<input type="checkbox" name="kaupungit" value="tampere"<% if instr(config, "tampere") > 0 or RS("config") = "" then response.write "checked" %>> tampere<br>
<input type="checkbox" name="kaupungit" value="turku"<% if instr(config, "tukru") > 0 or RS("config") = "" then response.write "checked" %>> turku<br>
<input type="checkbox" name="kaupungit" value="espoo"<% if instr(config, "espoo") > 0 or RS("config") = "" then response.write "checked" %>> espoo<br>
<input type="checkbox" name="kaupungit" value="jyv‰skyl‰"<% if instr(config, "jyv‰skyl‰") > 0 or RS("config") = "" then response.write "checked" %>> jyv‰skyl‰<br>
<input type="checkbox" name="kaupungit" value="kotka"<% if instr(config, "kotka") > 0 or RS("config") = "" then response.write "checked" %>> kotka<br>
<input type="checkbox" name="kaupungit" value="kouvola"<% if instr(config, "kouvola") > 0 or RS("config") = "" then response.write "checked" %>> kouvola
		</td>
		<td valign="top">
<br>
<input type="checkbox" name="kaupungit" value="kuopio"<% if instr(config, "kuopio") > 0 or RS("config") = "" then response.write "checked" %>> kuopio<br>
<input type="checkbox" name="kaupungit" value="lahti"<% if instr(config, "lahti") > 0 or RS("config") = "" then response.write "checked" %>> lahti<br>
<input type="checkbox" name="kaupungit" value="lappeenranta"<% if instr(config, "lappeenranta") > 0 or RS("config") = "" then response.write "checked" %>> lappeenranta<br>
<input type="checkbox" name="kaupungit" value="oulu"<% if instr(config, "oulu") > 0 or RS("config") = "" then response.write "checked" %>> oulu<br>
<input type="checkbox" name="kaupungit" value="pori"<% if instr(config, "pori") > 0 or RS("config") = "" then response.write "checked" %>> pori<br>
<input type="checkbox" name="kaupungit" value="vaasa"<% if instr(config, "vaasa") > 0 or RS("config") = "" then response.write "checked" %>> vaasa<br>
<input type="checkbox" name="kaupungit" value="-muu-"<% if instr(config, "-muu-") > 0 or RS("config") = "" then response.write "checked" %>> muut<br>
<input type="checkbox" name="kaupungit" value="kaikki" <% if instr(config, "kaikki") > 0 or RS("config") = "" then response.write "checked" %>> kaikki<br>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<td valign="top">
n‰yt‰ tiedot:<br>
<input type="checkbox" name="nayta" value="uusi" <% if instr(config, "uusi") > 0 or RS("config") = "" then response.write "checked" %>> uusi/muutettu<br>
<input type="checkbox" name="nayta" value="hinta" <% if instr(config, "hinta") > 0 or RS("config") = "" then response.write "checked" %>> hinta<br>
<input type="checkbox" name="nayta" value="ikaraja" <% if instr(config, "ikaraja") > 0 or RS("config") = "" then response.write "checked" %>> ik‰raja
	</td>
	<td valign="top">
<br>
<input type="checkbox" name="nayta" value="auki" <% if instr(config, "auki") > 0 or RS("config") = "" then response.write "checked" %>> aukioloaika<br>
<input type="checkbox" name="nayta" value="kaupunki" <% if instr(config, "kaupunki") > 0 or RS("config") = "" then response.write "checked" %>> kaupunki<br>
	</td>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td align="left" colspan="2">
<input type="hidden" name="muutos" value="true">
<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok">
		</td>
	</tr>
	</table>
	</td>
</tr>
</table>
</form>
<%
RS.Close
Con.Close
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