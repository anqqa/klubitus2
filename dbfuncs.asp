<%
response.Buffer = true
Session.LCID = 1035 'suomi-asetukset
dbPath = Server.MapPath("\klubitus\dbs\rekisteri.mdb")

adOpenStatic = 3
adLockOptimistic = 3

paivat = array("maanantai", "tiistai", "keskiviikko", "torstai", "perjantai", "lauantai", "sunnuntai")
kuut = array("TAMMIKUU", "HELMIKUU", "MAALISKUU", "HUHTIKUU", "TOUKOKUU", "KESƒKUU", "HEINƒKUU", "ELOKUU", "SYYSKUU", "LOKAKUU", "MARRASKUU", "JOULUKUU")

set Con = server.createobject("adodb.connection")

sub readyDBCon
	if Con = "" or isnull(Con) then
		set Con = Server.CreateObject("adodb.Connection")
		Con.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & dbPath & ";PWD=SALASANA"
'		Con.Open "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & dbPath '& ";PWD=SALASANA"
	end if
end sub

function formatURL(url)
	set temp = url
	if left(lcase(temp), 4) <> "http" then temp = "http://" & temp
	if (instr(temp, ".") < 1) or (instr(lcase(temp), "http://") < 1) then
		temp = " "
	end if
	formatURL = temp
end function

function emailOK(email)
	if (instr(email, ".") < 1) or (instr(email, "@") < 1) then
		emailOK = false
	else
		emailOK = true
	end if
end function

function zeroString(s)
	if s = "" then zeroString = " " else zeroString = s
end function

function unzeroString(s)
	if s = " " then unzeroString = "" else unzeroString = s
end function

sub mail(mailFrom, mailFromAddress, mailTo, mailToAddress, mailSubject, mailBody)
	Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
	Mailer.FromName = mailFrom
	Mailer.FromAddress = mailFromAddress
	Mailer.RemoteHost = "smtpgw.activeisp.com"
	Mailer.AddRecipient mailTo, MailToAddress
	Mailer.Subject = mailSubject
	Mailer.BodyText = mailBody
	Mailer.SendMail
end sub

function paiva(pvm)
	paiva = datepart("d", pvm) & "." & datepart("m", pvm) & "." & datepart("yyyy", pvm) & " / " & right(pvm, 11)
end function

function paiva2(pvm)
	paiva2 = datepart("d", pvm) & "." & datepart("m", pvm) & "." & datepart("yyyy", pvm)
end function

function zero(num)
	if num < 10 then temp = "0" & num else temp = num end if
	zero = temp
end function

function chat_parser(num, aika, nikki, viesti, ip, alias, aliasto, selain, block)
	viesti_alku = ""
	viesti_loppu = "<br>"
	viesti_viesti = ""
	viesti_aika = formatdatetime(dateadd("h", 1, aika(num)), 3)
	komento = ""
	ok = false
'komento
	if left(viesti(num), 1) = "/" then
'komento: /msg
		if lcase(left(viesti(num), 5)) = "/msg " then
			temp = right(viesti(num), len(viesti(num)) - 5)
			viesti_alkaa = instr(temp, " ") - 1
			if viesti_alkaa > 1 then
				nimi = lcase(left(temp, viesti_alkaa))
				temp = right(temp, len(temp) - viesti_alkaa)
				komento = "* <span class=sub2>" & viesti_aika & " [<b>" & nikki(num) & "</b>] -> (<b>" & nimi & "</b>) " & temp & "</span><br>"
			else
				komento = "* <i>tais unohtua ite viesti?</i><br>"
			end if
'			komento = "* privaatti<br>"
			ok = true
		end if
'komento: /me
		if not ok and lcase(left(viesti(num), 4)) = "/me " then
			viesti_alku = " ---- "
			viesti_nikki = " <b><i>" & nikki(num) & "</i></b> "
			viesti_viesti = right(viesti(num), len(viesti(num)) - 4)
			ok = true
		end if
'komento: /help
		if not ok and lcase(left(viesti(num), 5)) = "/help" or lcase(left(viesti(num), 5)) = "/apua" then
			komento = "* <b>komennot:</b><br>* /msg nimi viesti = <i>ei n‰‰ muut</i><br>* /me viesti = <i>toimintaa..</i><br>* /ip nimi = <i>n‰ytt‰‰ ip-osoitteen</i><br>* /whois nimi = <i>n‰ytt‰‰ tarkempaa infoa</i><br>* /last x = <i>n‰ytt‰‰ x viimeist‰ rivi‰</i><br>* /smiley = <i>n‰ytt‰‰ smiley-listan</i><br>* /help = <i>jaa mik‰ lie..</i><br>"
			ok = true
		end if
'komento: /ip
		if not ok and lcase(left(viesti(num), 3)) = "/ip" then
			if len(viesti(num)) > 5 then nimi = lcase(right(viesti(num), len(viesti(num)) - 4)) else nimi = session("chat_nikki")
			loyty = false
			for a = 0 to 49
				if not loyty and lcase(nikki(a)) = nimi then
					loyty = true
					ipnum = ip(a)
				end if
			next
			komento = "* <i>" & nimi
			if loyty then
				komento = komento & ", ip " & ipnum & "</i><br>"
			else
				komento = komento & " ei oo paikalla..</i></br>"
			end if
			ok = true
		end if
'komento: /seeall
		if not ok and lcase(left(viesti(num), 7)) = "/seeall" then
			if session("admin") = "true" then
				if lcase(right(viesti(num), 2)) = "on" then 
					komento = "* see all: on<br>"
					session("chat_seeall") = "true" 
				else 
					komento = "* see all: off<br>"
					session("chat_seeall") = "false"
				end if
				ok = true
			end if
		end if
'komento: /alias
		if not ok and lcase(left(viesti(num), 6)) = "/alias" then
			komento = "* alias: "
			pituus = len(viestit(num))
			if pituus = 6 then
				for a = 0 to 19
					if alias(a) <> "" then komento = komento & left(alias(a), len(alias(a)) - 1) & " " & right(alias(a), 1) & "=" & aliasto(a) & ", "
				next
			else
				alias_to = right(viesti(num), len(viesti(num)) - 7)
				mihin = instr(alias_to, " ") - 1
				if mihin > 1 then
					alias_from = left(alias_to, mihin)
					alias_to = right(alias_to, len(alias_to) - mihin - 1)
					loyty = false
					for a = 0 to 19
						if not loyty and alias(a) = alias_from then
							loyty = true
							aliasto(a) = alias_to
						end if
					next
					if not loyty then
						for a = 0 to 19
							if not loyty and alias(a) = "" then 
								alias(a) = alias_from
								aliasto(a) = alias_to
								loyty = true
							end if
						next
					end if
					komento = komento & alias_from & " = " & alias_to
				else
					alias_from = right(viesti(num), len(viesti(num)) - 7)
					for a = 0 to 19
						if alias(a) = alias_from or alias_from = "kaikki" then
							alias(a) = ""
							aliasto(a) = ""
						end if
					next
					komento = komento & "poistettu " & alias_from
				end if
			end if
			if session("admin") <> "true" then komento = "* kiva peli."
			komento = komento & "<br>"
			ok = true
		end if
'komento: /smile
		if not ok and lcase(left(viesti(num), 6)) = "/smile" then
			komento = "* <table border=0 cellpadding=0 cellspacing=5><tr><td>:x-</td><td>: x -</td><td>:D</td><td>: D</td><td>B)</td><td>B )</td><td>;)D</td><td>; ) D</td><td>:o.</td><td>: o .</td><td>;)</td><td>; )</td><td>8(</td><td>8 (</td><td>;(</td><td>; (</td></tr><tr><td>=D</td><td>= D</td><td>:P</td><td>: P</td><td>d:)</td><td>d : )</td><td>:o)</td><td>: o )</td><td>:/</td><td>: /</td><td>:pb</td><td>: p b</td><td>x(</td><td>x (</td><td>:B</td><td>: B</td></tr><tr><td>3-o</td><td>3 - o</td><td>:s</td><td>: s</td><td>:x</td><td>: x</td><td>x)</td><td>x )</td><td>o:)</td><td>o : )</td><td>x:(</td><td>x : (</td><td>:Q</td><td>: Q</td><td>:)B</td><td>: ) B</td></tr><tr><td>:oo</td><td>: o o</td><td>:.(</td><td>: . (</td><td>:)</td><td>: )</td><td>:(</td><td>: (</td><td>s:o</td><td>s : o</td><td>;P</td><td>; P</td><td>:o</td><td>: o</td><td>=)</td><td>= )</td></tr><tr><td>:x)</td><td>: x )</td><td>;D</td><td>; D</td><td>8o</td><td>8 o</td><td>=/</td><td>= /</td><td>:666</td><td>: 6 6 6</td><td>-o/</td><td>- o /</td></tr></table>"
			ok = true
		end if
'komento: /shake
		if not ok and lcase(left(viesti(num), 6)) = "/shake" then
			if len(viesti(num)) > 7 then nimi = lcase(right(viesti(num), len(viesti(num)) - 7)) else nimi = session("chat_nikki")
			if nikki(num) = "anqqa" then komento = "* <i>zzzap " & nimi & "</i><br>"
			ok = true
		end if
'komento: /whois
		if not ok and lcase(left(viesti(num), 6)) = "/whois" then
			if len(viesti(num)) > 7 then nimi = lcase(right(viesti(num), len(viesti(num)) - 7)) else nimi = lcase(session("chat_nikki"))
			loyty = false
			for a = 49 to 0 step -1
				if not loyty and lcase(nikki(a)) = nimi then
					loyty = true
					ipnum = ip(a)
					selaintiedot = selain(a)
				end if
			next
			if loyty then
				set Con = Server.CreateObject("adodb.Connection")
				Con.Open "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & dbPath & ";PWD=SALASANA"
				set RS = Server.CreateObject("ADODB.Recordset")
				RS.ActiveConnection = Con
				RS.CursorType = adOpenStatic
				RS.Open "SELECT * FROM members WHERE nick ='" & nimi & "'"
				if not (RS.BOF and RS.EOF) then
					klubitustiedot = "<i>" & nimi & " (#" & RS("id") & ")</i>:<br>"
					if len(RS("city")) > 2 then	klubitustiedot = klubitustiedot & "kaupunki: " & RS("city") & "<br>"
					if len(RS("homepage")) > 10 then klubitustiedot = klubitustiedot & "kotisivu: <a href=" & formaturl(RS("homepage")) & " target=_new>" & RS("homepage") & "</a><br>"
					klubitustiedot = klubitustiedot & "rekisterˆitynyt: " & RS("join") & "<br>"
					klubitustiedot = klubitustiedot & "k‰yntej‰: " & RS("visits") & "<br>"
					klubitustiedot = klubitustiedot & "viestej‰: " & RS("messages") & "<br>"
					klubitustiedot = klubitustiedot & "bilelis‰yksi‰: " & RS("adds") & " (" & RS("mods") & ")<br>"
				else
					klubitustiedot = "<i>" & nimi & " (vierailija)</i>:<br>" 
				end if
				RS.Close
				set RS = nothing
				Con.Close
				Set HttpObj = Server.CreateObject("AspHTTP.Conn")
				HttpObj.UserAgent = "Mozilla/4.0 (compatible; MSIE 5.0; Windows NT 4.0)"
				url = "http://whois.tulevaisuus.net/ketaon.php?query=" & ipnum
				HttpObj.Url = url
				HttpObj.TimeOut = 20
				strResult = replace(HttpObj.GetURL, "&nbsp;", " ")
				tulos = instr(strResult, "class=""querylinkki"">")
				if tulos > 0 then
					tulos = right(strResult, len(strResult) - tulos - 19)
					tulos = left(tulos, instr(tulos, "</a>") - 1)
					if instr(strResult, "RIPE Whois") then
						netname = right(strResult, len(strResult) - instr(strResult, "netname:") + 1)
						netname = left(netname, instr(netname, "country:") - 1)
						netname = replace(netname, vbnewline, "")
						netname = replace(netname, vblf, "")
					end if
				else
					tulos = "ei saada yhteytt‰.."
					netname = ""
				end if
				komento = "* " & klubitustiedot & "ip " & ipnum & "</i> = " & tulos & "<br>" & netname & "software: " & selaintiedot & "<br><a href=http://whois.tulevaisuus.net/ketaon.php?query=" & ipnum & " target=_new>tarkempi info</a><br>"
			else
				komento = "* <i>" & nimi & "</i> ei oo paikalla..<br>"
			end if
			ok = true
		end if
'komento: /block
		if not ok and lcase(left(viesti(num), 6)) = "/block" and session("admin") = "true" then
			if len(viesti(num)) > 7 then 
				nimi = lcase(right(viesti(num), len(viesti(num)) - 7))
				loyty = false
				vapaa = -1
				for a = 0 to 19
					if not loyty and block(a) = nimi then 
						block(a) = ""
						loyty = true
					end if
					if not loyty and vapaa = -1 and block(a) = "" then vapaa = a
				next
				if not loyty then 
					block(vapaa) = nimi
					komento = "* <i>" & nimi & "</i> blockattu, paikka " & vapaa & "..<br>"
				else
					komento = "* <i>" & nimi & "</i> poistettu blockista..<br>"
				end if
			else
				komento = "* blockissa: "
				for a = 0 to 20
					if block(a) <> "" then komento = komento & block(a) & " "
				next
				komento = komento & "<br>"
			end if
			ok = true
		end if
		if not ok then
			komento = "* mik‰ ihmeen " & viesti(num) & ".. ?<br>"
		end if
'komento: /anqqa
'		if not ok and lcase(left(viesti(num), 6)) = "/anqqa" then
'			if len(viesti(num)) > 7 then nimi = lcase(right(viesti(num), len(viesti(num)) - 7)) else nimi = session("chat_nikki")
'			if nikki(num) = "anqqa" then komento = "* <i>kvaak " & nimi & "</i><br>" else komento = "* ei kvaaki<br>"
'			ok = true
'		end if
'normaali
	else
		viesti_nikki = " [<b>" & nikki(num) & "</b>]"
		viesti_viesti = " " & viesti(num)
	end if
	if komento = "" then 
		chat_parser = viesti_alku & viesti_aika & viesti_nikki & viesti_viesti & viesti_loppu 
	else
		chat_parser = komento
	end if
end function

function linjoilla(ip)
	dim porukkaa_aika, porukkaa_ip, porukkaa_nikki
	yht = 0
	vierasta = 0
	nimia = ""
	if application("porukkaa_ok") <> "true" then
		redim porukkaa_aika(100), porukkaa_ip(100), porukkaa_nikki(100)
		porukkaa_aika(0) = now()
		porukkaa_ip(0) = ip
		if session("nikki") = "" then
			porukkaa_nikki(0) = "vieras"
		else
			porukkaa_nikki(0) = session("nikki")
		end if
		yht = 1
	else
		porukkaa_aika = application("porukkaa_aika")
		porukkaa_ip = application("porukkaa_ip")
		porukkaa_nikki = application("porukkaa_nikki")
		up = false
		for p = 0 to 100
			if porukkaa_ip(p) = ip then
				porukkaa_aika(p) = now()
				yht = yht + 1
				up = true
			elseif porukkaa_aika(p) <> "" then
				if datediff("n", porukkaa_aika(p), now()) > 5 then
					porukkaa_aika(p) = ""
					porukkaa_ip(p) = ""
					porukkaa_nikki(p) = ""
				else
					yht = yht + 1
				end if
			end if
		next
		if not up then
			for p = 0 to 100
				if porukkaa_ip(p) = "" then
					porukkaa_ip(p) = ip
					porukkaa_aika(p) = now()
					if session("nikki") = "" then
						porukkaa_nikki(p) = "vieras"
					else
						porukkaa_nikki(p) = session("nikki")
					end if
					yht = yht + 1
					exit for
				end if
			next
		end if
	end if
	for p = 0 to 100
		if porukkaa_nikki(p) = "vieras" then
			vierasta = vierasta + 1
		elseif porukkaa_nikki(p) <> "" then
			nimia = nimia & porukkaa_nikki(p) & ", "
		end if
	next
	nimia = nimia & vierasta & " vierasta"
	application.lock
	application("porukkaa_ok") = "true"
	application("porukkaa_aika") = porukkaa_aika
	application("porukkaa_ip") = porukkaa_ip
	application("porukkaa_nikki") = porukkaa_nikki
	application.unlock
	linjoilla = yht
	nimet = nimia
end function
%>
