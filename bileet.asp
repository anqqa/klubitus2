<!-- #INCLUDE FILE="dbfuncs.asp" -->
<html>
<head>
	<title>bileet</title>
<link rel="stylesheet" href="liila.css" type="text/css">
<%
kk = request.querystring("kk")
v = request.querystring("v")
d = request.querystring("d")
etsi = request.form("etsi")
if (kk = "") then kk = month(date())
if (v = "") then v = datepart("yyyy", date())
if (d = "") then d = 0
kk = int(kk) + int(d)
if (kk > 12) then 
	kk = 1
	v = int(v) + 1
end if
if (kk < 1) then
	kk = 12
	v = int(v) - 1
end if

function viikonpaiva(p, kk, v)
	temp = weekday(p & "/" & kk & "/" & v, vbmonday)
	viikonpaiva = left(paivat(temp - 1), 2)
end function

dim tanaan, rivi
rivi = "<tr §tanaan§>" & vbNewline
rivi = rivi & "<td valign='top'>§uus§</td>" & vbNewline
rivi = rivi & "<td valign='top'>§muokattu§</td>" & vbNewline
rivi = rivi & "<td valign='top' class='paiva'><nobr>§paiva§</nobr></td>" & vbNewline
rivi = rivi & "<td valign='top' class='sub1'>§pileet§<br><span class='sub2'>§paikka§</span></td>" & vbNewline
rivi = rivi & "<td valign='top'>dj: §dj§<br>§muuta§§erotin2§§musiikki§</td>" & vbNewline
rivi = rivi & "<td valign='top'><nobr>§hinta§§erotin§§ikaraja§<br>§auki§</nobr></td>" & vbNewline
rivi = rivi & "<td valign='top' align='right'>§kaupunki§</td>" & vbNewline
rivi = rivi & "</tr>" & vbNewline
if instr(session("config"), "uusi") = 0 and session("config") <> "" then
	rivi = replace(rivi, "§uus§", "")
	rivi = replace(rivi, "§muokattu§", "")
end if
if instr(session("config"), "hinta") = 0 and session("config") <> "" then rivi = replace(rivi, "§hinta§", "")
if (instr(session("config"), "hinta") = 0 and instr(session("config"), "ikaraja") = 0) and session("config") <> "" then rivi = replace(rivi, "§erotin§", "")
if instr(session("config"), "ikaraja") = 0 and session("config") <> "" then rivi = replace(rivi, "§ikaraja§", "")
if instr(session("config"), "kaupunki") = 0 and session("config") <> "" then rivi = replace(rivi, "§kaupunki§", "")
if instr(session("config"), "auki") = 0 and session("config") <> "" then rivi = replace(rivi, "§auki§", "")

sub bile(num)
	rivi2 = rivi
	if session("kaupunki") = "" or session("kaupunki") = bileet(5, num) then
	if instr(session("config"), bileet(5, num)) or instr(session("config"), "kaikki") or session("config") = "" then
		if bileet(0, num) = datepart("d", date()) and kk = datepart("m", date()) then
			if tanaan = "" then 
				response.write "<a name=""tanaan"">" & vbNewline
				tanaan = "true"
			end if
			rivi2 = replace(rivi2, "§tanaan§", "bgcolor='#990099'")
		else
			rivi2 = replace(rivi2, "§tanaan§", "")
		end if
		if dateDiff("d", bileet(13, num), Now()) < 4 then 
			rivi2 = replace(rivi2, "§uus§", "+") 
		else 
			rivi2 = replace(rivi2, "§uus§", "")
		end if
		if dateDiff("d", bileet(14, num), Now()) < 4 then 
			rivi2 = replace(rivi2, "§muokattu§", "¤")
		else
			rivi2 = replace(rivi2, "§muokattu§", "")
		end if
		if session("nikki") <> "" then 
			rivi2 = replace(rivi2, "§paiva§", "<a href='bileet_muokkaa.asp?id=" & bileet(15, num) & "' class='paiva'>" & viikonpaiva(bileet(0, num), kk, v) & " " & bileet(0, num) & "</a>.")
		else
			rivi2 = replace(rivi2, "§paiva§", viikonpaiva(bileet(0, num), kk, v) & " " & bileet(0, num) & ".")
		end if
		if bileet(2, num) = " " then
			rivi2 = replace(rivi2, "§pileet§", bileet(1, num))
		else
			rivi2 = replace(rivi2, "§pileet§", "<a href='" & bileet(2, num) & "' class='sub1' target='_blank'>" & bileet(1, num) & "</a>")
		end if
		if bileet(10, num) > -1 then 
			rivi2 = replace(rivi2, "§hinta§", bileet(10, num) & "mk")
		else
			rivi2 = replace(rivi2, "§hinta§", "")
		end if
		if bileet(10, num) > -1 and bileet(9, num) > 0 then 
			rivi2 = replace(rivi2, "§erotin§", ", ")
		else
			rivi2 = replace(rivi2, "§erotin§", "")
		end if
		if bileet(9, num) > 0 then 
			rivi2 = replace(rivi2, "§ikaraja§", "K" & bileet(9, num))
		else
			rivi2 = replace(rivi2, "§ikaraja§", "")
		end if
		rivi2 = replace(rivi2, "§dj§", bileet(6, num))
		rivi2 = replace(rivi2, "§kaupunki§", bileet(5, num))
		if bileet(4, num) = " " then
			rivi2 = replace(rivi2, "§paikka§", bileet(3, num))
		else
			rivi2 = replace(rivi2, "§paikka§", "<a href='" & bileet(4, num) & "' class='sub2' target='_blank'>" & bileet(3, num) & "</a>")
		end if
		auki = ""
		if bileet(7, num) > -1 then auki = zero(bileet(7, num)) & "-" 
		if bileet(8, num) > -1 then auki = auki & zero(bileet(8, num))
		rivi2 = replace(rivi2, "§auki§", auki)
		rivi2 = replace(rivi2, "§muuta§", bileet(12, num))
		if bileet(12, num) <> " " and bileet(11, num) <> " " then 
			rivi2 = replace(rivi2, "§erotin2§", " - ")
		else
			rivi2 = replace(rivi2, "§erotin2§", "")
		end if
		rivi2 = replace(rivi2, "§musiikki§", bileet(11, num))
		response.write rivi2
	end if
	end if
end sub

readyDBCon
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
if etsi <> "" then
	SQL = "SELECT day, name, name_url, place, place_url, city, dj, hours_from, hours_to, age, price, music, info, added, modified, id, month, year FROM parties WHERE name LIKE '%" & etsi & "%' OR place LIKE '%" & etsi & "%' OR dj LIKE '%" & etsi & "%' OR music LIKE '%" & etsi & "%' OR info LIKE '%" & etsi & "%' ORDER BY year, month, day"
else
	SQL = "SELECT day, name, name_url, place, place_url, city, dj, hours_from, hours_to, age, price, music, info, added, modified, id FROM parties WHERE month = " & kk & " AND year = " & v & " ORDER BY day, city, name"
end if
RS.Open SQL
yht = int(RS.RecordCount)
if yht > 0 then bileet = RS.GetRows() ' kenttä, numero
RS.Close
Con.Close
%>
</head>

<body bgcolor="#000000" background="pics/tausta_liila.jpg" bgproperties="fixed">

<% if etsi = "" then %>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td align="left"><a href="bileet.asp?kk=<%=kk %>&v=<%=v %>&d=-1" class="navi">edellinen</a></td>
	<td align="center" valign="bottom">

<form method="post" action="bileet.asp">
<table cellpadding="0" cellspacing="0" border="0">
<tr>
	<td><input type="text" name="etsi" size="20" value="">&nbsp;</td>
	<td><input style="background: none" type="image" name="submit" src="pics/icon_etsi.gif" alt="etsi" width="12" height="22" hspace="5" align="absmiddle" border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
<select name="pika" onChange="window.location.href=this.options[this.selectedIndex].value">
<%
	for vuodet = 2001 to 2002
		for kuukaudet = 1 to 12
			if (vuodet & kuukaudet = v & kk) then
				response.write "<option value=""bileet.asp?kk=" & kuukaudet & "&v=" & vuodet & """ selected>" & lcase(kuut(kuukaudet - 1)) & ", " & vuodet & "</option>" & vbNewline
			else
				response.write "<option value=""bileet.asp?kk=" & kuukaudet & "&v=" & vuodet & """>" & lcase(kuut(kuukaudet - 1)) & ", " & vuodet & "</option>" & vbNewline
			end if
		next 
	next
	response.flush
%>
</select>
	</td>
</tr>
</table>
</form>

	</td>
	<td align="right"><a href="bileet.asp?kk=<%=kk %>&v=<%=v %>&d=+1" class="navi">seuraava</a></td>
</tr>
<tr>
	<td class="kuukausi" colspan="2"><%=kuut(kk - 1) %>, <%=v %></td>
	<td align="right">+ = uus, ¤ = muokattu
<%
	if session("admin") = "true" then response.write "· <a href='bileet_hae.asp' class='navi'>hae bileet</a>"
	if session("nikki") <> "" then response.write "· <a href='bileet_lisaa.asp?kk=" & kk & "&v=" & v &"' class='navi'>uus bile</a>"
%>
	</td>
</tr>
<tr>
	<td colspan="3">
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<%
	if yht > 0 then
		for i = 0 to yht - 1
			bile(i)
			if i / 10 = int(i / 10) then response.flush
		next
	end if
%>
</table>	

	</td>
</tr>
<tr>
	<td align="left" colspan="2"><a href="bileet.asp?kk=<%=kk %>&v=<%=v %>&d=-1" class="navi">edellinen</a></td>
	<td align="right"><a href="bileet.asp?kk=<%=kk %>&v=<%=v %>&d=+1" class="navi">seuraava</a></td>
</tr>
</table>
<% else %>
<table width="100%" cellpadding="5" cellspacing="0" border="0">
<tr>
	<td align="center" colspan="7">

<form method="post" action="bileet.asp">
<table cellpadding="0" cellspacing="0" border="0">
<tr>
	<td><input type="text" name="etsi" size="20" value="<%=etsi %>">&nbsp;</td>
	<td><input style="background: none" type="image" name="submit" src="pics/icon_etsi.gif" alt="etsi" width="12" height="22" hspace="5" align="absmiddle" border="0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<td>
<select name="pika" onChange="window.location.href=this.options[this.selectedIndex].value">
<%
	for vuodet = 2001 to 2002
		for kuukaudet = 1 to 12
			response.write "<option value='bileet.asp?kk=" & kuukaudet & "&v=" & vuodet & "'>" & lcase(kuut(kuukaudet - 1)) & ", " & vuodet & "</option>"
		next 
	next
%>
</select>
	</td>
</tr>
</table>
</form>

	</td>
</tr>
<tr>
	<td colspan="5">löyty yhteensä <%=yht %></td>
	<td align="right" colspan="2">+ = uus, ¤ = muokattu
<%
	if session("nikki") <> "" then response.write "· <a href='bileet_lisaa.asp?kk=" & kk & "&v=" & v &"' class='navi'>uus bile</a>"
%>
	</td>
</tr>
<%
	if yht > 0 then 
		kk = bileet(16, 0)
		v = bileet(17, 0)
%>
<tr>
	<td class="kuukausi" colspan="7"><%=kuut(kk - 1) %>, <%=v %></td>
</tr>
<%
		for i = 0 to yht - 1
			if kk <> bileet(16, i) or v <> bileet(17, i) then
				kk = bileet(16, i)
				v = bileet(17, i)
%>
<tr>
	<td class="kuukausi" colspan="7"><%=kuut(kk - 1) %>, <%=v %></td>
</tr>
<%
				bile(i)
			else
				bile(i)
			end if
		next
%>
</table>
<% 
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
