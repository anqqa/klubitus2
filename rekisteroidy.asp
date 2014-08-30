<html>
<head>
	<title>rekisteroidy</title>
	<link rel="stylesheet" href="keltainen.css" type="text/css">
</head>

<body bgcolor="#000000" background="pics/tausta_keltainen.jpg" bgproperties="fixed">

<form action="rekkaus.asp" method="post">
<table cellspacing="0" cellpadding="0" border="0">
<tr>
	<td class="selitys" colspan="2">
<b>huom!</b><br>
<ul>
<li>jos oot unohtanu salasanan, täytä <u>vain</u> käyttäjätunnus ja salasana tulee meilissä
<li><b>lihavoidut</b> kohdat on pakollisia
<li><i>kursivoidut</i> kohdat näkyy muille, muut ei
</ul>
<br>
	</td>
</tr>
<tr>
	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td class="selite"><b><i>käyttäjätunnus:</i></b></td>
		<td><input type="text" maxlength="20" size="30" name="nikki"></td>
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
		<td><textarea name="kuvaus" rows="5" cols="30"></textarea></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite" valign="top"><i>signature:</i></td>
		<td><textarea name="signature" cols="30"></textarea></td>
	</tr>
	</table>
	</td>

	<td valign="top" align="left" width="50%">
	<table cellspacing="5" cellpadding="0" width="100%" border="0">
	<tr>
		<td class="selite"><b>nimi:</b></td>
		<td><input type="text" maxlength="30" size="30" name="nimi"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><b>e-mail:</b></td>
		<td><input type="text" maxlength="60" size="30" name="email"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><i>kotisivu:</i></td>
		<td><input type="text" maxlength="60" size="30" name="kotisivu" value="http://"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite"><i>kaupunki:</i></td>
		<td><input type="text" maxlength="30" size="30" name="kaupunki"></td>
	</tr>
	<tr>
		<td colspan="2">&nbsp;</td>
	</tr>
	<tr>
		<td class="selite">syntymäaika:</td>
		<td>
		<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td>
<select name="synt_paiva">
<option value="0" selected>--</option>
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
<option value="20">30.</option>
<option value="31">31.</option>
</select>
			</td>
			<td>
<select name="synt_kuukausi">
<option value="0" selected>--</option>
<option value="1">tammikuuta</option>
<option value="2">helmikuuta</option>
<option value="3">maaliskuuta</option>
<option value="4">huhtikuuta</option>
<option value="5">toukokuuta</option>
<option value="6">kesäkuuta</option>
<option value="7">heinäkuuta</option>
<option value="8">elokuuta</option>
<option value="9">syyskuuta</option>
<option value="10">lokakuuta</option>
<option value="11">marraskuuta</option>
<option value="12">joulukuuta</option>
</select>
			</td>
			<td>
<select name="synt_vuosi">
<option value="0" selected>--</option>
<option value="1990">1990</option>
<option value="1989">1989</option>
<option value="1988">1988</option>
<option value="1987">1987</option>
<option value="1986">1986</option>
<option value="1985">1985</option>
<option value="1984">1984</option>
<option value="1983">1983</option>
<option value="1982">1982</option>
<option value="1981">1981</option>
<option value="1980">1980</option>
<option value="1979">1979</option>
<option value="1978">1978</option>
<option value="1977">1977</option>
<option value="1976">1976</option>
<option value="1975">1975</option>
<option value="1974">1974</option>
<option value="1973">1973</option>
<option value="1972">1972</option>
<option value="1971">1971</option>
<option value="1970">1970</option>
<option value="1969">1969</option>
<option value="1968">1968</option>
<option value="1967">1967</option>
<option value="1966">1966</option>
<option value="1965">1965</option>
<option value="1964">1964</option>
<option value="1963">1963</option>
<option value="1962">1962</option>
<option value="1961">1961</option>
<option value="1960">1960</option>
<option value="1959">1959</option>
<option value="1958">1958</option>
<option value="1957">1957</option>
<option value="1956">1956</option>
<option value="1955">1955</option>
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
		<td class="selite" colspan="2">
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