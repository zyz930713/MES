<%
function getSubPart(isshowname,showstyle,subpart,where,order,sequency,splitchar,SUB_ID,SUB_Number)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
rcount=0
set rsSP=server.CreateObject("adodb.recordset")
SQLSP="Select 1,P.NID,P.PART_NUMBER,P.SECTION_ID,S.SECTION_NAME from Part P inner join SECTION S on P.SECTION_ID=S.NID inner join FACTORY F on P.FACTORY_ID=F.NID where P.ROUTINE_TYPE<>0"&where&order
rsSP.open SQLSP,conn,1,3
if not rsSP.eof then
	if sequency<>"" then
		aseq=split(sequency,",")
		for x=0 to ubound(aseq)
			while not rsSP.eof
				if aseq(x)=rsSP("NID") then
					select case showstyle
					case "OPTION"
					output=output&"<option value='"&rsSP("NID")&"'"
						if rsSP("NID")=part then
						output=output&"selected"
						end if
					output=output&">"&rsSP("PART_NUMBER")&"</option>"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsSP("PART_NUMBER")&splitchar
						else
						output=output&rsSP("PART_NUMBER")
						end if
					end select
				end if
			rsSP.movenext
			wend
		rsSP.movefirst
		next
	else
		y=1
		tabi=1
		while not rsSP.eof
			select case showstyle
			case "OPTION"
				output=output&"<option value='"&rsSP("NID")&"'"
					if rsSP("NID")=part then
					output=output&"selected"
					end if
				output=output&">"&rsSP("PART_NUMBER")&" ("&rsSP("SECTION_NAME")&")</option>"
			case "TEXT"
				if y<>rcount then
				output=output&rsSP("PART_NUMBER")&splitchar
				else
				output=output&rsSP("PART_NUMBER")
				end if
			case "ARRAY"
				SUB_ID=SUB_ID&rsSP("NID")&","
				SUB_Number=SUB_Number&rsSP("PART_NUMBER")&","
			end select
		rsSP.movenext
		y=y+1
		wend
	end if
end if
getSubPart=output
	if SUB_ID<>"" then
	SUB_ID=left(SUB_ID,len(SUB_ID)-1)
	SUB_Number=left(SUB_Number,len(SUB_Number)-1)
	end if
rsSP.close
set rsSP=nothing
end function
%>