<%
function getOperator(isname,showstyle,code,where,order,sequence,splitchar,opcount)
opcount=0
	'isstationname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
	
set rsO=server.CreateObject("adodb.recordset")
SQLO="Select O.CODE from OPERATORS O "&where&order
rsO.open SQLO,conn,1,3
if not rsO.eof then
opcount=rsO.recordcount
end if
rsO.close
SQLO="Select 1,O.CODE,O.OPERATOR_NAME,O.OPERATOR_CHINESE_NAME from OPERATORS O left join FACTORY F on O.FACTORY_ID=F.NID"&where&order
rsO.open SQLO,conn,1,3
if not rsO.eof then
	if sequence<>"" then
		aseq=split(sequence,",")
		for x=0 to ubound(aseq)
			while not rsO.eof
				if aseq(x)=rsO("CODE") then
					select case showstyle
					case "OPTION"
					output=output&"<option value='"&rsO("CODE")&"'"
						if rsO("CODE")=code then
						output=output&"selected"
						end if
					output=output&">"&rsO("CODE")
						if isname=true and rsO("OPERATOR_NAME")<>"" then
						output=output&" ("&rsO("OPERATOR_NAME")&" "&rsO("OPERATOR_CHINESE_NAME")&")"
						end if
					output=output&"</option>"
					case "TEXT"
						output=output&rsO("OPERATOR_NAME")&splitchar
					end select
				end if
			rsO.movenext
			wend
		rsO.movefirst
		next
	else
		while not rsO.eof
			select case showstyle
			case "OPTION"
			output=output&"<option value='"&rsO("CODE")&"'"
				if rsO("CODE")=code then
				output=output&"selected"
				end if
			output=output&">"&rsO("CODE")
				if isname=true and rsO("OPERATOR_NAME")<>"" then
				output=output&" ("&rsO("OPERATOR_NAME")&" "&rsO("OPERATOR_CHINESE_NAME")&")"
				end if
				output=output&"</option>"
			case "TEXT"
			output=output&rsO("OPERATOR_NAME")&splitchar
			end select
		rsO.movenext
		wend
	end if
end if
getOperator=output
rsO.close
set rsO=nothing
end function
%>