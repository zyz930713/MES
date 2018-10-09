<%
function getSeriesGroup(showstyle,seriesgroup,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select S.NID,S.SERIES_GROUP_NAME from SERIES_GROUP S inner join FACTORY F on S.FACTORY_ID=F.NID "&where&order
session("aerror")=SQLS
'response.End()
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=seriesgroup then
			output=output&"selected"
			end if
		output=output&">"&rsS("SERIES_GROUP_NAME")&"</option>"
		case "CHARTOPTION"
		output=output&"<option value='"&rsS("SERIES_GROUP_NAME")&"'"
			if rsS("SERIES_GROUP_NAME")=seriesgroup then
			output=output&"selected"
			end if
		output=output&">"&rsS("SERIES_GROUP_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='seriesgroup' type='radio' class='t-c-greenCopy' value='"&rsS("NID")&"'>"&rsS("SERIES_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsS("SERIES_GROUP_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getSeriesGroup=output
rsS.close
set rsS=nothing
end function

function getFinanceSeriesGroup(showstyle,seriesgroup,where,order,splitchar)
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select S.NID,S.SERIES_GROUP_NAME from FINANCE_SERIES_GROUP S inner join FACTORY F on S.FACTORY_ID=F.NID "&where&order
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=seriesgroup then
			output=output&"selected"
			end if
		output=output&">"&rsS("SERIES_GROUP_NAME")&"</option>"
		case "CHARTOPTION"
		output=output&"<option value='"&rsS("SERIES_GROUP_NAME")&"'"
			if rsS("SERIES_GROUP_NAME")=seriesgroup then
			output=output&"selected"
			end if
		output=output&">"&rsS("SERIES_GROUP_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='seriesgroup' type='radio' class='t-c-greenCopy' value='"&rsS("NID")&"'>"&rsS("SERIES_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsS("SERIES_GROUP_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getFinanceSeriesGroup=output
rsS.close
set rsS=nothing
end function
%>