<?xml version="1.0"?> 
<!DOCTYPE wml PUBLIC "-//WAPFORUM//DTD WML 1.1//EN" "http://www.wapforum.org/DTD/wml_1.1.xml"> 

<%
p = request.querystring("p")
kk = request.querystring("kk")
v = request.querystring("v")
kaupunki = request.querystring("kaupunki")
%> 
<wml> 
 
  <card id="bileet" title="$(kaupunki) $(p).$(kk).$(v)"> 
    <do type="prev" label="hae"><go href="#hae"/></do> 
    <p><small>
<%
readyDBcon
set RS = Server.CreateObject("ADODB.Recordset")
RS.ActiveConnection = Con
RS.CursorType = adOpenStatic
SQL = "SELECT id, name, place FROM parties WHERE day = " & p & " AND month = " & kk & " AND year = " & v & " AND city = " & kaupunki & " ORDER BY name"
RS.Open SQL
yht = int(RS.RecordCount)
if yht > 0 then bileet = RS.GetRows()
RS.Close
Con.Close
if yht > 0 then 
	for i = 1 to yht - 1
		response.write "<a href=""bile.wml?id=" & bileet(0, i) & """>" & bileet(1, num) & "</a> @ " & bileet(2, num) & "<br/>"
	next
end if
%>
    </small></p> 
  </card> 
   
  <card id="hae" title="hae p&#xE4;iv&#xE4;"> 
    <do type="prev" label="hae"><go href="#bileet"/></do> 
    <p>
p&#xE4;iv&#xE4;: <input type="text" format="* N" name="p" maxlength="2" value="$(p)"/>
kuukausi: <input type="text" format="* N" name="kk" maxlength="2" value="$(kk)"/>
vuosi: <input type="text" format="NNNN" name="v" value="$(v)"/>
kaupunki:
<select name="kaupunki" value="$(kaupunki)" title="kaupunki">
<option value="helsinki">helsinki</option>
<option value="järvenpää">j&#xE4;rvenp&#xE4;&#xE4;</option>
<option value="tampere">tampere</option>
<option value="turku">turku</option>
<option value="espoo">espoo</option>
<option value="jyväskylä">jyv&#xE4;skyl&#xE4;</option>
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

    </p> 
  </card> 
 
</wml> 
