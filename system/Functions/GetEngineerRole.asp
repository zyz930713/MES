<%
function getEngineerRole(isshowname,showstyle,role,where,order,sequency,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsE=server.CreateObject("adodb.recordset")
SQLE="Select NID,ROLE_NAME from ENGINEER_ROLE "& where&order
rsE.open SQLE,conn,1,3
if not rsE.eof then
z=rsE.recordcount
	if sequency<>"" then
		aseq=split(sequency,",")
		for x=0 to ubound(aseq)
			while not rsE.eof
				if aseq(x)=rsE("NID") then
					select case showstyle
					case "OPTION"
					output=output&"<option value='"&rsE("ROLE_NAME")&"'"
						if rsE("NID")=role then
						output=output&"selected"
						end if
					output=output&">"&rsE("ROLE_NAME")&"</option>"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsE("ROLE_NAME")&splitchar
						else
						output=output&rsE("ROLE_NAME")
						end if
					end select
				end if
			rsE.movenext
			wend
		rsE.movefirst
		next
	else
		y=1
		while not rsE.eof
			select case showstyle
			case "OPTION"
				output=output&"<option value='"&rsE("ROLE_NAME")&"'"
					if rsE("NID")=role then
					output=output&"selected"
					end if
				output=output&">"&rsE("ROLE_NAME")&"</option>"
			case "TEXT"
				if y<>z then
				output=output&rsE("ROLE_NAME")&splitchar
				else
				output=output&rsE("ROLE_NAME")
				end if
			end select
		rsE.movenext
		y=y+1
		wend
	end if
end if
getEngineerRole=output
rsE.close
set rsE=nothing
end function
%>