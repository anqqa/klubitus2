<!--#include file="lib.asp"-->
<%
if session("admin") = "" then response.redirect "virhe.asp"
paikka = request.querystring("paikka")

if paikka <> "" then
	openDB
	set RS = Server.CreateObject("ADODB.Recordset")
	RS.ActiveConnection = Con
	RS.CursorType = adOpenStatic
	RS.Open "SELECT * FROM places WHERE id = " & paikka
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>paikat . muokkaa</title>
	<link rel="stylesheet" href="sin.css" type="text/css">
</head>

<body background="gfx/tausta_sin.gif" bgproperties="fixed">

<table cellpadding="0" cellspacing="5" border="0">
<form action="paikat_tallenna.asp" method="post">
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
<option value="j�rvenp��">j�rvenp��</option>
<option value="tampere">tampere</option>
<option value="turku">turku</option>
<option value="espoo">espoo</option>
<option value="jyv�skyl�">jyv�skyl�</option>
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
	<td class="selite">ik�raja:</td><td><input type="text" name="ikaraja" size="2"></td>
</tr>
<tr>
	<td align="right" colspan="2">
<input style="background:transparent" type="image" height="36" alt="ok" width="28" src="pics/icon_ok.gif" border="0" name="ok">&nbsp;<a href="javascript:history.go(-1)"><img height="36" alt="antaa olla" src="pics/icon_pyh.gif" width="28" border="0"></a>
	</td>
</tr>
</form>
</table>

</body>
</html>
