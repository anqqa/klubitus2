<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = -1000 

readyDBCon

set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT id, topic, entered, message FROM messages WHERE status = 11 ORDER BY entered DESC"
RS.Open SQL
yht = int(RS.RecordCount)
if yht > 0 then uutiset = RS.GetRows()
RS.Close
Con.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>uutiset</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<table cellspacing="0" cellpadding="3" border="0">
<tr>
	<td colspan="2">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
	<td align="left" class="head">uutiset</td>
	<td align="right">
uutisia yhteens‰: 
<%
response.write yht
if session("admin") = "true" then response.write " ∑ <a href=""uutiset_lisaa.asp"" class=""sub2"">lis‰‰</a> ∑ <a href=""newsletter.asp?ok=true"" class=""sub2"">news letter</a>"
%>
	</td>
</tr>
</table>
	</td>
</tr>
<%
if yht > 0 then
	for i = 0 to yht - 1
		response.write "<tr bgcolor=""#009900"">" & vbNewline
		response.write "<td><b>" & uutiset(1, i) & "</b></td>" & vbNewline
		response.write "<td class=""sub1"" align=""right""><nobr>" & uutiset(2, i) & "</nobr></td>" & vbNewline
		response.write "</tr>" & vbNewline
		response.write "<tr>" & vbNewline
		response.write "<td colspan=""2"">" & replace(uutiset(3, i), vbNewline, "<br>") & "</td>" & vbNewline
		response.write "</tr>" & vbNewline
		response.write "<tr>" & vbNewline
		response.write "<td colspan=""2"">&nbsp;</td>" & vbNewline
		response.write "</tr>" & vbNewline
	next
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
