<%
haku = int(request.querystring("haku"))
Session.LCID = 1035 'suomi-asetukset
dbPath = Server.MapPath("\klubitus\dbs\rekisteri.mdb")
adOpenStatic = 3
adLockOptimistic = 3
paivat = array("maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai", "sunnuntai")

tanaan = date()
dim viikko(8)
for i = 0 to 7
	viikko(i) = dateadd("d", i, tanaan)
next

set Con = Server.CreateObject( "adodb.Connection" )
Con.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & dbPath & ";PWD=SALASANA"
set RS = Server.CreateObject( "ADODB.Recordset" )
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT day, name, place, city, dj, hours_from, hours_to, age, price, music, info, day, month, year FROM parties ORDER BY year, month, day, city, name"
RS.Open SQL
yht = int(RS.RecordCount)
bileet = RS.GetRows() ' kenttä, numero
RS.Close
Con.Close
d = haku
if haku > 0 then
	for p = 1 to 7
		if weekday(viikko(p), vbmonday) = haku then d = p
	next
end if
response.write day(viikko(d)) & "." & month(viikko(d)) & ".:" & vbnewline
for i = 0 to yht - 1
	pvm = datevalue(bileet(11, i) & "/" & bileet(12, i) & "/" & bileet(13, i))
	if pvm = viikko(d) then
		response.write bileet(1, i) & "@" & bileet(2, i) & vbnewline
		yht2 = yht2 + 1
	end if
next
if yht2 = 0 then response.write "ei löydy bileitä haetulle päivälle.." & vbnewline
response.write sisalto
%>