<%
haku = request.querystring("haku")
sanat = split(lcase(haku), " ")
sanoja = ubound(sanat) + 1

set objFSO = CreateObject("Scripting.FileSystemObject") 
set laskuri = objFSO.OpenTextFile(server.mappath("smslaskuri.txt"), 1, false)
latauksia = int(laskuri.readline)
laskuri.close
latauksia = latauksia + 1
set laskuri = objFSO.OpenTextFile(server.mappath("smslaskuri.txt"), 2, true)
laskuri.writeline latauksia
laskuri.close
set laskuri = objFSO.OpenTextFile(server.mappath("smshaut.txt"), 8, false)
laskuri.writeline now() & ": " & haku
laskuri.close
set laskuri = nothing
set objFSO = nothing

moodi = ""
if haku = "" then
	moodi = "apua"
end if
if moodi = "" then
' viikonpäivä?
	kaupunki = "helsinki"
	tanaan = date()
	paivat = array("tänään", "maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai", "sunnuntai")
	paiva = -1
	for p = 0 to 7
		if sanat(0) = paivat(p) or sanat(0) = left(paivat(p), 2) then paiva = p
	next
	if paiva <> -1 then
		dim viikko(8)
		hakupaiva = -1
		for i = 0 to 7
			viikko(i) = dateadd("d", i, tanaan)
			if weekday(viikko(i), vbmonday) = paiva then hakupaiva = viikko(i)
		next
		if paiva = 0 then hakupaiva = viikko(0)
		moodi = "päivä"
		if sanoja > 1 then kaupunki = sanat(1)
	end if
end if
' paikka?
if moodi = "" then 
	if sanat(0) = "paikka" then
		moodi = "paikka"
		if sanoja > 1 then paikka = sanat(1) else paikka = "eioo"
	end if
end if
' päivämäärä?
if moodi = "" then
	moodi = "päivämäärä"
	if sanoja > 1 then
	' päivämäärä ja kaupunki?
		if instr(sanat(0), ".") > 0 then 
			hp = split(sanat(0), ".")
			kaupunki = sanat(1)
			moodi = "päivämääräkaupunki"
		end if
	' päivämäärä ja bileet?
		if instr(sanat(1), ".") > 0 then 
			hp = split(sanat(1), ".")
			bile = sanat(0)
			moodi = "päivämääräbileet"
		end if
	else
		hp = split(sanat(0), ".")
	end if
	select case ubound(hp)
		case 2
			hakupaiva = datevalue(hp(0) & "/" & hp(1) & "/" & year(tanaan))
		case 3
			hakupaiva = datevalue(hp(0) & "/" & hp(1) & "/" & hp(2))
		case else
			moodi = ""
	end select
end if
' bileet?
if moodi = "" then 
	moodi = "bileet"
	bile = sanat(0)
end if

if moodi <> "apua" then
	Session.LCID = 1035 'suomi-asetukset
	dbPath = Server.MapPath("\klubitus\dbs\rekisteri.mdb")
	adOpenStatic = 3
	adLockOptimistic = 3
	set Con = Server.CreateObject( "adodb.Connection" )
	Con.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & dbPath & ";PWD=SALASANA"
	set RS = Server.CreateObject( "ADODB.Recordset" )
	RS.ActiveConnection = Con
	RS.CursorType = adOpenStatic
	if moodi = "paikka" then
		SQL = "SELECT * FROM places ORDER BY name"
	else
		SQL = "SELECT day, name, place, city, dj, hours_from, hours_to, age, price, music, info, day, month, year FROM parties ORDER BY year, month, day, city, name"
	end if
	RS.Open SQL
	yht = int(RS.RecordCount)
	if yht > 0 then	bileet = RS.GetRows() ' kenttä, numero
	RS.Close
	Con.Close
end if

select case moodi
	case "päivä", "päivämäärä", "päivämääräkaupunki"
		if moodi = "päivämäärä" then response.write left(paivat(weekday(hakupaiva, vbmonday)), 2) & " "
		response.write day(hakupaiva) & "." & month(hakupaiva) & ".:\n"
		yht2 = 0
		if yht > 0 then
			for i = 0 to yht - 1
				pvm = datevalue(bileet(11, i) & "/" & bileet(12, i) & "/" & bileet(13, i))
				if pvm = hakupaiva and bileet(3, i) = kaupunki then
					response.write bileet(1, i) & "@" & bileet(2, i) & "\n"
					yht2 = yht2 + 1
				end if
			next
		else
			response.write "ei löydy bileitä haetulle päivälle.."
		end if
		if yht2 = 0 then response.write "ei löydy bileitä haetulle päivälle.."
	case "päivämääräbileet"
		response.write bile & " " & left(paivat(weekday(hakupaiva, vbmonday)), 2) & " " & day(hakupaiva) & "." & month(hakupaiva) & ". "
		yht2 = 0
		if yht > 0 then
			for i = 0 to yht - 1
				pvm = datevalue(bileet(11, i) & "/" & bileet(12, i) & "/" & bileet(13, i))
				if pvm = hakupaiva and instr(bileet(1, i), bile) > 0 then
					if bileet(5, i) <> "" then
						if bileet(6, i) <> "" then 
							response.write zero(bileet(5, i)) & "-" & zero(bileet(6, i)) & " "
						else
							response.write zero(bileet(5, i)) & "-> "
						end if
					else
						if bileet(6, i) <> "" then response.write "->" & bileet(6, i) & " "
					end if
					if bileet(7, i) > 0 then response.write "k" & bileet(7, i) & " "
					if bileet(8, i) > -1 then response.write bileet(8, i) & "mk "
					if len(bileet(4, i)) > 1 then response.write "\ndj: " & bileet(4, i)
					yht2 = yht2 + 1
				end if
			next
		else
			response.write "\nei löydy bilettä haetulle päivälle.."
		end if
		if yht2 = 0 then response.write "\nei löydy bilettä haetulle päivälle.."
	case "bileet"
		yht2 = 0
		if yht > 0 then
			for i = 0 to yht - 1
				pvm = datevalue(bileet(11, i) & "/" & bileet(12, i) & "/" & bileet(13, i))
				if datediff("d", tanaan, pvm) > -1 and instr(bileet(1, i), bile) > 0 and yht2 = 0 then
					response.write bileet(1, i) & " " & left(paivat(weekday(pvm, vbmonday)), 2) & " " & day(pvm) & "." & month(pvm) & ". "
					if bileet(5, i) > -1 then
						if bileet(6, i) > -1 then 
							response.write zero(bileet(5, i)) & "-" & zero(bileet(6, i)) & " "
						else
							response.write zero(bileet(5, i)) & "-> "
						end if
					else
						if bileet(6, i) > -1 then response.write "->" & bileet(6, i) & " "
					end if
					if bileet(7, i) > 0 then response.write "k" & bileet(7, i) & " "
					if bileet(8, i) > -1 then response.write bileet(8, i) & "mk "
					if len(bileet(4, i)) > 1 then response.write "\ndj: " & bileet(4, i)
					yht2 = yht2 + 1
				end if
			next
		else
			response.write "\nei löydy bilettä.."
		end if
		if yht2 = 0 then response.write bile & ": " & "\nei löydy bilettä.."
	case "paikka"
		yht2 = 0
		for i = 0 to yht - 1
			if instr(bileet(1, i), paikka) > 0 then
				response.write bileet(1, i) & "\n"
				if len(bileet(2, i)) > 10 then response.write right(bileet(2, i), len(bileet(2, i)) - 7) & "\n"
				if bileet(4, i) <> "" then response.write bileet(4, i) & "\n"
				if bileet(5, i) <> "" then response.write "puh. " & bileet(5, i) & "\n"
				if bileet(6, i) <> "" then response.write "puh. " & bileet(6, i) & "\n"
				if bileet(7, i) <> "" then response.write "fax " & bileet(7, i) & "\n"
				if bileet(9, i) > 0 then response.write "auki: " & bileet(9, i) & "-" & zero(bileet(10, i)) & "\n"
				if bileet(11, i) > 0 then response.write "ikäraja: " & bileet(11, i)
				yht2 = 1
			end if
		next
		if yht = 0 then response.write paikka & ": ei löydy klubia/ravintolaa.."
	case else
		response.write "ma/ti/ke/to/pe/la/su/tänään - päivän bileet\n24.10. - päivän bileet\n24.10. klubi - tarkemmat tiedot\npaikka ctrl/vanha/jne - ravintolan tiedot"
end	select

function zero(num)
	if num < 10 then temp = "0" & num else temp = num end if
	zero = temp
end function
%>