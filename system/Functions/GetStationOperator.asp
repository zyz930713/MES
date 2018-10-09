<%
function getStationOperator(jobnumber,sheetnumber,jobtype,station_id,repeated_sequence,sequency,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsO=server.CreateObject("adodb.recordset")
SQLO="Select 1,JS.STATION_ID,JS.REPEATED_SEQUENCE,JS.OPERATOR_CODE,O.PRACTISED from JOB_STATIONS JS left join OPERATORS O on JS.OPERATOR_CODE=O.CODE where JS.JOB_NUMBER='"&jobnumber&"' and JS.SHEET_NUMBER='"&sheetnumber&"' and JS.JOB_TYPE='"&jobtype&"' and JS.STATION_ID in ('"&station_id&"')"
rsO.open SQLO,conn,1,3
z=rsO.recordcount
if not rsO.eof then
	if sequency<>"" then
		aseq=split(sequency,",")
		if isnull(repeated_sequence) or repeated_sequence="" then
			for x=0 to ubound(aseq)
				repeated_sequence=repeated_sequence&"1,"
			next
			repeated_sequence=left(repeated_sequence,len(repeated_sequence)-1)
		end if
		arepeated=split(repeated_sequence,",")
		if ubound(aseq)<>ubound(arepeated) then
		diff=ubound(aseq)-ubound(arepeated)
		ReDim Preserve arepeated(UBound(arepeated) +diff)
		end if
		for x=0 to ubound(aseq)
			while not rsO.eof
				if aseq(x)=rsO("STATION_ID") and arepeated(x)=cstr(rsO("REPEATED_SEQUENCE")) then
					if x<>ubound(aseq) then
						if rsO("PRACTISED")="1" then
						output=output&"<span class='t-c-practised'>"&rsO("OPERATOR_CODE")&"</span>"&splitchar
						else
						output=output&rsO("OPERATOR_CODE")&splitchar
						end if
					else
						if rsO("PRACTISED")="1" then
						output=output&"<span class='t-c-practised'>"&rsO("OPERATOR_CODE")&"</span>"
						else
						output=output&rsO("OPERATOR_CODE")
						end if
					end if
				end if
			rsO.movenext
			wend
		rsO.movefirst
		next
	else
		y=1
		while not rsO.eof
			if y<>z then
				if rsO("PRACTISED")="1" then
				output=output&"<span class='t-c-practised'>"&rsO("OPERATOR_CODE")&"</span>"&splitchar
				else
				output=output&rsO("OPERATOR_CODE")&splitchar
				end if
			else
				if rsO("PRACTISED")="1" then
				output=output&"<span class='t-c-practised'>"&rsO("OPERATOR_CODE")&"</span>"
				else
				output=output&rsO("OPERATOR_CODE")
				end if
			end if
		rsO.movenext
		y=y+1
		wend
	end if
end if
getStationOperator=output
rsO.close
set rsO=nothing
end function
%>