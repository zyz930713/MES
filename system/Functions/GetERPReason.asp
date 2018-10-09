<%
function getERPReason()
set rsS=server.CreateObject("adodb.recordset")
SQLD="Select * from REMOTE_TRANSACTION_REASON where reason_id in (8, 130, 132) order by REASON_NAME"
rsS.open SQLD,conn,1,3
while not rsS.eof 
	optiontext=optiontext&"<option value='"&rsS("REASON_ID")&"'>"&rsS("REASON_NAME")&" ("&left(rsS("DESCRIPTION"),20)&")</option>"
rsS.movenext
wend
getERPReason=optiontext
rsS.close
set rsS=nothing
end function
%>