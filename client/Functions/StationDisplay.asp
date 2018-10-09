<%
function display_stations(astation,atransaction,display_stations_id)
for j=1 to ubound(astation)
	if atransaction(j)="0" or atransaction(j)="2" then
	display_stations_id=display_stations_id&","&astation(j)
	exit for
	end if
next
display_stations=display_stations_id
end function

function stations_refresh(astation,atransaction,last_station_id,stations_index,stations_transaction)
new_stations_index=""
new_stations_transaction=""
for j=0 to ubound(astation)'get new stations index
	new_stations_index=new_stations_index&astation(j)&","
	new_stations_transaction=new_stations_transaction&atransaction(j)&","
	if astation(j)=last_station_id then
	exit for
	end if
next
stations_index=left(new_stations_index,len(new_stations_index)-1)
stations_transaction=left(new_stations_transaction,len(new_stations_transaction)-1)
end function

function last_for_conj(astation,atransaction)
for j=ubound(astation) to lbound(astation) step -1 'get the last station
	if atransaction(j)="2" then
	last_for_conj=astation(j)
	exit for
	end if
next
end function

function last_for_notconj(astation,atransaction)
for j=ubound(astation) to lbound(astation) step -1 'get the last station
	if atransaction(j)="0" then
	last_for_notconj=astation(j)
	exit for
	end if
next
end function
%>