<%
function getStationTransactionChange(where,sequency)
	'where=where phrase of SQL query
output=" "
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select NID,TRANSACTION_CHANGE from STATION "&where
session("aerror")=SQLS
rsS.open SQLS,conn,1,3
if not rsS.eof then
	aseq=split(sequency,",")
	for x=0 to ubound(aseq)
		while not rsS.eof
			if aseq(x)=rsS("NID") then
				output=output&rsS("TRANSACTION_CHANGE")&","
			end if
		rsS.movenext
		wend
	rsS.movefirst
	next
end if
getStationTransactionChange=left(output,len(output)-1)
rsS.close
set rsS=nothing
end function
%>