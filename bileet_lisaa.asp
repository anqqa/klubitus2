<html>
<head>
	<title>bileet . lis‰‰</title>
	<link rel="stylesheet" href="liila.css" type="text/css">
<%
if session("nikki") = "" then response.redirect "virhe.asp"
kk = request.querystring("kk")
v = request.querystring("v")
if (kk = "") then kk = month(date())
kk = int(kk)
if (v = "") then v = year(date())
v = int(v)
%>
</head>

<body bgcolor="#000000" background="pics/tausta_liila.jpg" bgproperties="fixed">

<form action="bileet_lisays.asp" method="post">
<table cellspacing="0" cellpadding="0" border="0">
<tr>
	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td class="selite"><b>p‰iv‰m‰‰r‰:</b></td>
		<td>
<select name="p">
<option value="1">1.</option>
<option value="2">2.</option>
<option value="3">3.</option>
<option value="4">4.</option>
<option value="5">5.</option>
<option value="6">6.</option>
<option value="7">7.</option>
<option value="8">8.</option>
<option value="9">9.</option>
<option value="10">10.</option>
<option value="11">11.</option>
<option value="12">12.</option>
<option value="13">13.</option>
<option value="14">14.</option>
<option value="15">15.</option>
<option value="16">16.</option>
<option value="17">17.</option>
<option value="18">18.</option>
<option value="19">19.</option>
<option value="20">20.</option>
<option value="21">21.</option>
<option value="22">22.</option>
<option value="23">23.</option>
<option value="24">24.</option>
<option value="25">25.</option>
<option value="26">26.</option>
<option value="27">27.</option>
<option value="28">28.</option>
<option value="29">29.</option>
<option value="30">30.</option>
<option value="31">31.</option>
</select> 

<select name="kk">
<option value="1" <%if kk = 1 then response.write("selected") %>>tammikuuta</option>
<option value="2" <%if kk = 2 then response.write("selected") %>>helmikuuta</option>
<option value="3" <%if kk = 3 then response.write("selected") %>>maaliskuuta</option>
<option value="4" <%if kk = 4 then response.write("selected") %>>huhtikuuta</option>
<option value="5" <%if kk = 5 then response.write("selected") %>>toukokuuta</option>
<option value="6" <%if kk = 6 then response.write("selected") %>>kes‰kuuta</option>
<option value="7" <%if kk = 7 then response.write("selected") %>>hein‰kuuta</option>
<option value="8" <%if kk = 8 then response.write("selected") %>>elokuuta</option>
<option value="9" <%if kk = 9 then response.write("selected") %>>syyskuuta</option>
<option value="10" <%if kk = 10 then response.write("selected") %>>lokakuuta</option>
<option value="11" <%if kk = 11 then response.write("selected") %>>marraskuuta</option>
<option value="12" <%if kk = 12 then response.write("selected") %>>joulukuuta</option>
</select>

<select name="v">
<option value="2000" <%if v = 2000 then response.write("selected") %>>2000</option>
<option value="2001" <%if v = 2001 then response.write("selected") %>>2001</option>
<option value="2002" <%if v = 2002 then response.write("selected") %>>2002</option>
<option value="2003" <%if v = 2003 then response.write("selected") %>>2003</option>
</select>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><b>nimi:</b></td>
		<td><input type="text" maxlength="100" size="30" value="" name="nimi"></td>
	</tr>
	<tr>
		<td class="selite"> - kotisivu:</td>
		<td><input type="text" maxlength="100" size="30" value="http://" name="nimi_url"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite" valign="top">paikka:</td>
		<td><input type="text" maxlength="100" size="30" value="" name="paikka"></td>
	</tr>
	<tr>
		<td class="selite" valign="top"> - kotisivu:</td>
		<td><input type="text" maxlength="100" size="30" value="http://" name="paikka_url"></td>
	</tr>
	<tr>
		<td class="selite">ik‰raja:</td>
		<td>K<input type="text" maxlength="2" size="1" name="ikaraja"></td>
	</tr>
	<tr>
		<td class="selite">auki:</td>
		<td>
		<table cellpadding="0" cellspacing="0" border="0">
		<tr>
		<td>
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
		</td>
		<td>&nbsp;-&nbsp;</td>
		<td>
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
		</table>
		</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite" valign="top"><b>kaupunki:</b></td>
		<td>
<select name="kaupunki">
<option value="helsinki">helsinki</option>
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
</select>
		</td>
	</tr>
	</table>
	</td>

	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td class="selite">dj:</td>
		<td><input type="text" maxlength="250" size="30" name="dj"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">musiikkityyli:</td>
		<td><input type="text" maxlength="60" size="30" name="musiikki"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">muuta infoa:</td>
		<td><textarea name="muuta" cols="30" rows="3" wrap="virtual"></textarea></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">normaali hinta:</td>
		<td><input type="text" maxlength="3" size="1" name="hinta">mk</td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td align="right" colspan="2">
<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok">&nbsp;<a href="javascript:history.go(-1)"><img height="36" alt="antaa olla" src="pics/icon_pyh.gif" width="28" border="0"></a>
		</td>
	</tr>
	</table>
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