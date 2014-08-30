<!-- #INCLUDE FILE="dbfuncs.asp" -->
<%
Response.Expires = -1000 
'if session("nikki") = "" then response.redirect "virhe.asp"

alue = request.querystring("alue")
viesti = request.querystring("viesti")

readyDBCon
'set RS = Con.Execute("SELECT topic FROM messages WHERE id = " & alue)
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
if viesti <> "" then
	SQL = "SELECT a.topic, m.topic FROM message_areas AS a, messages AS m WHERE a.id = " & alue & " AND m.id = " & viesti
else
	SQL = "SELECT topic FROM message_areas WHERE id = " & alue
end if
RS.Open SQL
otsikko = RS.GetRows()
RS.Close
if session("nikki") <> "" then
	SQL = "SELECT signature FROM members WHERE id = " & session("id")
	RS.Open SQL
	sig = RS.GetRows()
	RS.Close
end if
Con.Close
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>forum . viesti . lis‰‰</title>
	<link rel="stylesheet" href="vihrea.css" type="text/css">
</head>

<body bgcolor="#002200" background="pics/tausta_vihrea.jpg" bgproperties="fixed">

<form action="forum_viesti_lisays.asp" method="post">
<table cellspacing="5" cellpadding="0" border="0" >
<tr>
	<td class="selite">alue:</td>
	<td class="sub2"><%=otsikko(0, 0) %></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<% if session("nikki") = "" then %>
<tr>
	<td class="selite">nimi:</td>
	<td class="sub2"><input type="text" maxlength="50" size="50" value="" name="nimi"></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<% end if 
if viesti <> "" then
	response.write "<input type=""hidden"" name=""aihe"" value=""" & otsikko(1, 0) & """>"
else
%>
<tr>
	<td class="selite">aihe:</td>
	<td><input type="text" maxlength="100" size="50" value="" name="aihe"></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<% end if %>
<tr>
	<td class="selite" valign="top">viesti:</td>
	<td><textarea cols="50" rows="20" name="sisalto"></textarea></td>
</tr>
<tr>
	<td colspan="2">&nbsp;</td>
</tr>
<tr>
	<td align="right" colspan="2">
<input type="hidden" name="alue" value="<%=alue %>">
<input type="hidden" name="viesti" value="<%=viesti %>">
<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok">&nbsp;<a href="javascript:history.go(-1)"><img height="36" alt="antaa olla" src="pics/icon_pyh.gif" width="28" border="0"></a>
	</td>
</tr>
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
