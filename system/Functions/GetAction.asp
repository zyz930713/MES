
<%
function getAction(isstationname,showstyle,action,where,order,sequence,splitchar)
	'isstationname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
	
set rsA=server.CreateObject("adodb.recordset")
SQLA="Select 1,A.NID,A.STATUS,A.ACTION_NAME,A.ACTION_CHINESE_NAME,S.STATION_NAME,F.FACTORY_NAME from ACTION A left join STATION S on A.STATION_ID=S.NID left join FACTORY F on A.FACTORY_ID=F.NID "&where&order
rsA.open SQLA,conn,1,3
if not rsA.eof then
	if sequence<>"" then
		aseq=split(sequence,",")
		for x=0 to ubound(aseq)
			while not rsA.eof
				if aseq(x)=rsA("NID") then
					select case showstyle
					case "OPTION"
					output=output&"<option value='"&rsA("NID")&"'"
						if rsA("NID")=action then
						output=output&"selected"
						end if
					output=output&">"&rsA("ACTION_NAME")&" ("&rsA("ACTION_CHINESE_NAME")&")"
						if isstationname=true and rsA("STATION_NAME")<>"" then
						output=output&" - "&rsA("STATION_NAME")&" - "&rsA("FACTORY_NAME")
						end if
					output=output&"</option>"
					case "TEXT"
						if rsA("STATUS")="0" then
						output=output&"<span class='dotline' title='status is disabled'>"&rsA("ACTION_NAME")&"</span>"&splitchar
						else
						output=output&"<span title='status is enabled'>"&rsA("ACTION_NAME")&"</span>"&splitchar
						end if
					end select
				end if
			rsA.movenext
			wend
		rsA.movefirst
		next
	else
		while not rsA.eof
			select case showstyle
			case "OPTION"
			output=output&"<option value='"&rsA("NID")&"'"
				if rsA("NID")=action then
				output=output&"selected"
				end if
			output=output&">"&rsA("ACTION_NAME")&" ("&rsA("ACTION_CHINESE_NAME")&")"
				if isstationname=true and rsA("STATION_NAME")<>"" then
				output=output&" - "&rsA("STATION_NAME")&" - "&rsA("FACTORY_NAME")
				end if
				output=output&"</option>"
			case "TEXT"
				if rsA("STATUS")="0" then
				output=output&"<span class='dotline' title='status is disabled'>"&rsA("ACTION_NAME")&"</span>"&splitchar
				else
				output=output&"<span title='status is enabled'>"&rsA("ACTION_NAME")&"</span>"&splitchar
				end if
			end select
		rsA.movenext
		wend
	end if
end if
getAction=output
rsA.close
set rsA=nothing
end function

session("getAction_New.output")=""
function getAction_New(isstationname,showstyle,action,where,order,sequence,splitchar)
output = session("getAction_New.output")
if output = "" then 
	set rsA=server.CreateObject("adodb.recordset")
	'order=" order by action_name"
	SQLA="Select 1,A.NID,A.ACTION_NAME,A.ACTION_CODE,A.ACTION_CHINESE_NAME,DECODE(STATION_POSITION,'0','Before','After') AS POSITION,F.FACTORY_NAME from ACTION_New A  left join FACTORY F on A.FACTORY_ID=F.NID where IS_Delete<>1 "&where&order
	rsA.open SQLA,conn,1,3
	if not rsA.eof then
		if sequence<>"" then
			aseq=split(sequence,",")
			for x=0 to ubound(aseq)
				while not rsA.eof
					if aseq(x)=rsA("NID") then
						select case showstyle
						case "OPTION"
						output=output&"<option value='"&rsA("NID")&"'"
							if rsA("NID")=action then
							output=output&"selected"
							end if
						output=output&">"&rsA("ACTION_NAME")&" ("&rsA("ACTION_CHINESE_NAME")&")"
							if isstationname=true  then
							output=output&" - "&rsA("FACTORY_NAME")
							end if
						output=output&"</option>"
						
						end select
					end if
				rsA.movenext
				wend
			rsA.movefirst
			next
		else
			while not rsA.eof
				select case showstyle
				case "OPTION"
					output=output&"<option value='"&rsA("NID")&"' >"
					output=output&rsA("ACTION_CODE")&" - "&rsA("ACTION_NAME")&" ("&rsA("ACTION_CHINESE_NAME")&")"&" - "&rsA("POSITION")
					if isstationname=true  then
						output=output&" - "&rsA("FACTORY_NAME")
					end if
					output=output&"</option>"				
				end select
			rsA.movenext
			wend
		end if
	end if
	
	rsA.close
	set rsA=nothing
	session("getAction_New.output") = output
end if
if action <> "" then
	output = replace(output,action&"'",action&"' selected")
end if
getAction_New=output
end function


%>