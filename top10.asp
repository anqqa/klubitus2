<!-- #INCLUDE FILE="dbfuncs.asp" -->
<% 
readyDBCon
set RS = Server.CreateObject( "ADODB.Recordset" )
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
RS.MaxRecords = 10
RS.Open "SELECT nick, visits, adds, mods, messages, id, id FROM members WHERE id > 1"
yht = int(RS.RecordCount)
lista = RS.GetRows() ' kenttä, numero
RS.Close
Con.Close

sub swap(a, b)
	c = a
	a = b
	b = c
end sub

function sort10(arr, field)
	dim anum
	redim anum(yht)
	for a = 0 to yht - 1
		anum(a) = a
	next
	for a = 0 to 9
		for b = a to yht - 1
			if arr(field, anum(b)) > arr(field, anum(a)) then swap anum(a), anum(b)
		next
	next
	sort10 = anum
end function

sub top10(tyyppi)
	dim temp, tauluheader, taulufooter, taulurivi
	tauluheader = "<table width=""100%"" cellspacing=""5"" cellpadding=""0"" border=""0"">" & vbNewline
	tauluheader = tauluheader + "<tr><td colspan=""3"" class=""head"">§otsikko§</td></tr>" & vbNewline
	taulurivi = "<tr><td align=""right"" class=""sub1"">§sija§.</td><td>§nimi§</td><td class=""sub2"">§pisteet§</td></tr>" & vbNewline
	taulufooter = "</table>"
	select case tyyppi
		case "score"
			tauluheader = replace(tauluheader, "§otsikko§", "top 10")
			top = 5
			for i = 0 to yht - 1
				lista(5, i) = lista(2, i) * 150 + lista(3, i) * 50 + lista(4, i) * 25 + lista(1, i) * 5
			next
		case "adds"
			tauluheader = replace(tauluheader, "§otsikko§", "top 10 lisäykset")
			top = 2
		case "visits"
			tauluheader = replace(tauluheader, "§otsikko§", "top 10 käynnit")
			top = 1
		case "messages"
			tauluheader = replace(tauluheader, "§otsikko§", "top 10 viestit")
			top = 4
	end select
	temp = sort10(lista, top)
	response.write tauluheader
	for i = 1 to 10
		rivi = replace(taulurivi, "§sija§", i)
		rivi = replace(rivi, "§nimi§", lista(0, temp(i-1)) & " (#" & lista(6, temp(i-1)) & ")")
		rivi = replace(rivi, "§pisteet§", lista(top, temp(i-1)))
		response.write rivi
	next
	response.write taulufooter
end sub
%>

<html>
<head>
	<title>top 10</title>
	<link rel="stylesheet" href="sininen.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_sininen.jpg" bgproperties="fixed">
<table border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="50%"><% top10("score") %></td><td width="50%"><% top10("adds") %></td>
</tr>
<tr>
	<td><% top10("visits") %></td><td><% top10("messages") %></td>
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