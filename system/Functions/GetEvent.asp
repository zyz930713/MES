<%
function getEvent(showtype,events,where,idcount,splitchar)
	if showtype="CHECKBOX" then
	optiontext="<table border='0' cellpadding='0' cellspacing='1'>"
	else
	optiontext=""
	end if
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select NID,EVENT_NAME from EVENT "&where
rsP.open SQLP,conn,1,3
idcount=recordcount
while not rsP.eof 
	if showtype="CHECKBOX" then
	optiontext=optiontext&"<tr class='t-c-GrayLight'><td><input name='event_id' type='checkbox' value='"&rsP("NID")&"'"
		if events<>"" then
			aevents=split(events,",")
			for k=0 to ubound(aevents)
				if rsP("NID")=aevents(k) then
				optiontext=optiontext&" checked"
				end if
			next
		end if
	optiontext=optiontext&">"&rsP("EVENT_NAME")&"</td></tr>"
	else
	optiontext=optiontext&rsP("EVENT_NAME")&splitchar
	end if
rsP.movenext
wend
if showtype="CHECKBOX" then
optiontext=optiontext&"</table>"
end if
getEvent=optiontext
rsP.close
set rsP=nothing
end function
%>