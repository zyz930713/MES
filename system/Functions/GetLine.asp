<%
function getLine(showstyle,line,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsP=server.CreateObject("adodb.recordset")
set rsScrap=server.CreateObject("adodb.recordset")
order=" order by  L.LINE_NAME"
SQLP="Select L.NID,L.LINE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
i=0
if not rsP.eof then
	while not rsP.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsP("NID")&"'"
			if rsP("NID")=line then
			output=output&"selected"
			end if
		output=output&">"&rsP("LINE_NAME")&"</option>"
		case "OPTION_LINE_NAME"
		output=output&"<option value='"&rsP("LINE_NAME")&"'"
			if rsP("LINE_NAME")=line then
			output=output&"selected"
			end if
		output=output&">"&rsP("LINE_NAME")&"</option>"
		case "RADIO"
			output=output&"<input name='factory' type='radio' class='t-c-greenCopy' value='"&rsP("NID")&"'>"&rsP("LINE_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsP("LINE_NAME")&splitchar
		case "CHECKBOX"
		
		'output=output&"<input type=""checkbox"" name=""members"" value="""&rsE("USER_CODE")&""" checked>"&rsE("USER_NAME")
			output=output&"<input type='checkbox' name='line' value='"&rsP("LINE_NAME")&"' "
			if instr(line,rsP("LINE_NAME"))>0 then
				output=output&" checked "
			end if
		    if (i mod 15)=0 and i>=15 then
			output=output&" >"&rsP("LINE_NAME")&"<BR>"
			else
			output=output&" >"&rsP("LINE_NAME")&"&nbsp;"
			end if
		case "CHECKBOX_SCRAP"
			SQL_Scrap="Select * from USERS_LINE where USER_CODE='"&session("code")&"' and LINE_ID='"&rsP("NID")&"'"
			rsScrap.open SQL_Scrap,conn,1,3
			if not rsScrap.eof then
				scrap_line=true
			else
				scrap_line=false
			end if
			rsScrap.close
			if scrap_line=true then
			output=output&"<input type=""checkbox"" name=""scrap_line"" value="""&rsP("NID")&""" checked>"&rsP("LINE_NAME")&"&nbsp;"
			else
			output=output&"<input type=""checkbox"" name=""scrap_line"" value="""&rsP("NID")&""">"&rsP("LINE_NAME")&"&nbsp;"
			end if
			if outi=12 then
			output=output&"<br>"
			outi=1
			else
			outi=outi+1
			end if
		end select
	i=i+1	
	rsP.movenext
	wend
end if
getLine=output
rsP.close
set rsP=nothing
set rsScrap=nothing
end function
%>