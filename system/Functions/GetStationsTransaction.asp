<%
function getStationsTransaction(stations)
astations=split(stations,",")
output=""
set rsST=server.CreateObject("adodb.recordset")
SQLST="select NID,TRANSACTION_TYPE from STATION where NID in ('"&replace(stations,",","','")&"')"
rsST.open SQLST,conn,1,3
if not rsST.eof then
	for i=0 to ubound(astations)
		while not rsST.eof
			if rsST("NID")=astations(i) then
			output=output&rsST("TRANSACTION_TYPE")&","
			end if
		rsST.movenext
		wend
		rsST.movefirst
	next
end if
rsST.close
getStationsTransaction=left(output,len(output)-1)
set rsST=nothing
end function

function getStationsTransaction_New(stations)
astations=split(stations,",")
output=""
set rsST1=server.CreateObject("adodb.recordset")
SQLST="select NID,TRANSACTION_TYPE from STATION_New where NID in ('"&replace(stations,",","','")&"')"
rsST1.open SQLST,conn,1,3
if not rsST1.eof then
	for i=0 to ubound(astations)
		while not rsST1.eof
			if rsST1("NID")=astations(i) then
			output=output&rsST1("TRANSACTION_TYPE")&","
			end if
		rsST1.movenext
		wend
		rsST1.movefirst
	next
end if
rsST1.close
getStationsTransaction_New=left(output,len(output)-1)
set rsST1=nothing
end function
%>