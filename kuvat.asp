<% if not session("index") = true then response.redirect "default.asp?sivu=kuvat.asp" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>kuvat</title>
<link rel="stylesheet" href="sininen.css" type="text/css">
<script language="javascript">
<!--
function ikkuna(kuva, bileet, numero, yhteensa){
	bileet = bileet.replace(/ /gi, "+");
	eval("window.open('kuva.asp?kuva=" + kuva + "&bileet=" + bileet + "&num=" + numero + "&yht=" + yhteensa + "', 'kuva','toolbar=no,location=no,directories=no,menubar=no,scrollbars=no,status=no,resizable=0,width=650,height=690')");
}
//-->
</script>
<%
sivu = request.querystring("sivu")
osa = request.querystring("osa")
if (sivu = "") then sivu = 20
if (osa = "") then osa = 1

function bileet(nimi, pileet, yhteensa)
	response.write "<tr><td class='head' colspan='5'>" & nimi & "</td></tr>" & vbNewline
	osia = fix(yhteensa / 30) + 1
	mista = 0
	mihin = yhteensa
	if (osia > 1) then
		mista = (osa - 1) * 30
		mihin = osa * 30
		if (mihin > yhteensa) then mihin = yhteensa
		response.write "<tr><td colspan='5'>"
		response.write "<a href=""kuvat.asp?sivu=" & sivu & "&osa=1"">sivu 1</a> "
		for i = 2 to osia
			response.write " | <a href=""kuvat.asp?sivu=" & sivu & "&osa=" & i & """>sivu " & i & "</a> "
		next
		response.write "</td></tr>" & vbNewline
	end if
	sivulla = mihin - mista
	kor = cint(sivulla / 5)
	if (sivulla mod 5 > 0) and (sivulla mod 5 < 3) then kor = kor + 1
	for y = 1 to kor
		response.write "<tr>" & vbNewline
		for x = 1 to 5
			if (((y - 1) * 5 + x) + mista <= mihin) then
				response.write "<td width='20%'>"
				response.write "<a href=""javascript:ikkuna('" & pileet & "', '" & nimi & "', " & ((y - 1) * 5 + x + mista) & ", " & yhteensa & ")""><img src='photot/mini/" & pileet & ((y - 1) * 5 + x + mista) & ".gif' border='0'></a>"
				response.write "</td>" & vbNewline
			end if
		next
		response.write "</tr>" & vbNewline
	next
end function
%>
</head>

<body bgcolor="#000000" text="#FFFFFF" background="pics/tausta_sininen.jpg" bgproperties="fixed">

<form>
<table width="100%" cellpadding="0" cellspacing="5" border="0">
<tr>
	<td colspan="5" align="center">
<%
dim pileita(20, 3)
pileita(20, 1) = "ministry of sound (13.10.2001)"
pileita(20, 2) = "mos"
pileita(20, 3) = 37
pileita(19, 1) = "reactor (8.9.2001)"
pileita(19, 2) = "reactor"
pileita(19, 3) = 11
pileita(18, 1) = "millennium wishmaster (1.9.2001)"
pileita(18, 2) = "wishmaster"
pileita(18, 3) = 55
pileita(17, 1) = "lab-4 (19.8.2001)"
pileita(17, 2) = "lab4"
pileita(17, 3) = 11
pileita(16, 1) = "labyrinth summer sound system (22.-23.6.2001)"
pileita(16, 2) = "lsss"
pileita(16, 3) = 13
pileita(15, 1) = "labyrinth 2001 (24.3.2001)"
pileita(15, 2) = "l2001"
pileita(15, 3) = 242
pileita(14, 1) = "club labyrinth (10.3.2001)"
pileita(14, 2) = "labyrinth103"
pileita(14, 3) = 24
pileita(13, 1) = "überjatkot (4.3.2001)"
pileita(13, 2) = "uber"
pileita(13, 3) = 13
pileita(12, 1) = "vertigo (3.3.2001)"
pileita(12, 2) = "vertigo33"
pileita(12, 3) = 7
pileita(11, 1) = "elements 9 (11.2.2001)"
pileita(11, 2) = "elements9"
pileita(11, 3) = 10
pileita(10, 1) = "pussy (3.2.2001)"
pileita(10, 2) = "pussy"
pileita(10, 3) = 18
pileita(9, 1) = "club labyrinth (20.1.2001)"
pileita(9, 2) = "clublabyrinth"
pileita(9, 3) = 13
pileita(8, 1) = "hype + illusions (2000/2001)"
pileita(8, 2) = "hypeillusions"
pileita(8, 3) = 26
pileita(7, 1) = "ever (27.12.2000)"
pileita(7, 2) = "ever"
pileita(7, 3) = 12
pileita(6, 1) = "sundissential (16.12.2000)"
pileita(6, 2) = "sundissential"
pileita(6, 3) = 43
pileita(5, 1) = "elements 8 (25.11.2000)"
pileita(5, 2) = "elements8"
pileita(5, 3) = 50
pileita(4, 1) = "suuri naamiaisjuhla (10.11.2000)"
pileita(4, 2) = "naamiaiset"
pileita(4, 3) = 27
pileita(3, 1) = "elements 7 (4.11.2000)"
pileita(3, 2) = "elements"
pileita(3, 3) = 38
pileita(2, 1) = "labyrinth 2000 (1.4.2000)"
pileita(2, 2) = "gatecrasher"
pileita(2, 3) = 26
pileita(1, 1) = "armin van buuren"
pileita(1, 2) = "armin"
pileita(1, 3) = 4
response.write("<select name=""pika"" onChange=""window.location.href=this.options[this.selectedIndex].value"">")
for a = 20 to 1 step -1
	if (pileita(a, 1) = pileita(sivu, 1)) then
		response.write("<option value=""kuvat.asp?sivu=" & a & """ selected>" & pileita(a, 1) & " - " & pileita(a, 3) & " kuvaa</option>") & vbNewline
	else
		response.write("<option value=""kuvat.asp?sivu=" & a & """>" & pileita(a, 1) & " - " & pileita(a, 3) & " kuvaa</option>") & vbNewline
	end if
next
response.write("</select></td>" & vbNewline & "</tr>" & vbNewline)
bileet pileita(sivu, 1), pileita(sivu, 2), pileita(sivu, 3)
%>
</table>
</form>
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
