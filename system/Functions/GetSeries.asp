<%
function getSeries(showstyle,series,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select S.NID,S.SERIES_NAME from SERIES S inner join FACTORY F on S.FACTORY_ID=F.NID "&where&order
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=series then
			output=output&"selected"
			end if
		output=output&">"&rsS("SERIES_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='series' type='radio' class='t-c-greenCopy' value='"&rsS("NID")&"'>"&rsS("SERIES_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsS("SERIES_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getSeries=output
rsS.close
set rsS=nothing
end function

function getFinanceSeries(showstyle,financeseries,where,order,splitchar)
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select S.NID,S.SERIES_NAME,F.FACTORY_NAME from FINANCE_SERIES S inner join FACTORY F on S.FACTORY_ID=F.NID "&where&order
'response.Write(SQLS)
'response.End()
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=financeseries then
			output=output&"selected"
			end if
		output=output&">"&rsS("SERIES_NAME")&" ("&rsS("FACTORY_NAME")&")</option>"
		case "RADIO"
		output=output&"<input name='series' type='radio' class='t-c-greenCopy' value='"&rsS("NID")&"'>"&rsS("SERIES_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsS("SERIES_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getFinanceSeries=output
rsS.close
set rsS=nothing
end function
%>