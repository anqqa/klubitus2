<%
if not session("index") = true then response.redirect "default.asp?sivu=liikkeet.asp"
if session("admin") = "" then response.redirect "virhe.asp"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>paikat . lis‰‰</title>
	<link rel="stylesheet" href="liila.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_liila.jpg" bgproperties="fixed">

<form action="paikat_lisays.asp" method="post">
<table cellpadding="0" cellspacing="5" border="0">
<tr>
	<td class="selite">paikka:</td><td><input type="text" name="nimi" size="20"></td>
</tr>
<tr>
	<td class="selite">- kotisivu:</td><td><input type="text" name="kotisivu" size="50"></td>
</tr>
<tr>
	<td class="selite">osoite:</td><td><input type="text" name="osoite" size="50"></td>
</tr>
<tr>
	<td class="selite">puhelin:</td><td><input type="text" name="puhelin" size="20"></td>
</tr>
<tr>
	<td class="selite">puhelin 2:</td><td><input type="text" name="puhelin2" size="20"></td>
</tr>
<tr>
	<td class="selite">fax:</td><td><input type="text" name="fax" size="20"></td>
</tr>
<tr>
	<td class="selite">kaupunki:</td><td><select name="kaupunki">
<option value="helsinki" selected">helsinki</option>
<option value="j‰rvenp‰‰">j‰rvenp‰‰</option>
<option value="tampere">tampere</option>
<option value="turku">turku</option>
<option value="espoo">espoo</option>
<option value="jyv‰skyl‰">jyv‰skyl‰</option>
<option value="kotka">kotka</option>
<option value="kouvola">kouvola</option>
<option value="kuopio">kuopio</option>
<option value="lahti">lahti</option>
<option value="lappeenranta">lappeenranta</option>
<option value="oulu">oulu</option>
<option value="pori">pori</option>
<option value="vaasa">vaasa</option>
<option value="-muu-">-muu-</option>
</select></td>
</tr>
<tr>
	<td class="selite">auki:</td><td>
<select name="auki1">
<option value="-1">--</option>
<option value="23">23</option>
<option value="22">22</option>
<option value="21">21</option>
<option value="20">20</option>
<option value="19">19</option>
<option value="18">18</option>
<option value="17">17</option>
<option value="16">16</option>
<option value="15">15</option>
<option value="14">14</option>
<option value="13">13</option>
<option value="12">12</option>
<option value="11">11</option>
<option value="10">10</option>
<option value="09">09</option>
<option value="08">08</option>
<option value="07">07</option>
<option value="06">06</option>
<option value="05">05</option>
<option value="04">04</option>
<option value="03">03</option>
<option value="02">02</option>
<option value="01">01</option>
<option value="00">00</option>
</select>
 - 
<select name="auki2">
<option value="-1">--</option>
<option value="00">00</option>
<option value="01">01</option>
<option value="02">02</option>
<option value="03">03</option>
<option value="04">04</option>
<option value="05">05</option>
<option value="06">06</option>
<option value="07">07</option>
<option value="08">08</option>
<option value="09">09</option>
<option value="10">10</option>
<option value="11">11</option>
<option value="12">12</option>
<option value="13">13</option>
<option value="14">14</option>
<option value="15">15</option>
<option value="16">16</option>
<option value="17">17</option>
<option value="18">18</option>
<option value="19">19</option>
<option value="20">20</option>
<option value="21">21</option>
<option value="22">22</option>
<option value="23">23</option>
</select>
	</td>
</tr>
<tr>
	<td class="selite">ik‰raja:</td><td><input type="text" name="ikaraja" size="2"></td>
</tr>
	<tr>
		<td align="right" colspan="2">
<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok">&nbsp;<a href="javascript:history.go(-1)"><img height="36" alt="antaa olla" src="pics/icon_pyh.gif" width="28" border="0"></a>
		</td>
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
