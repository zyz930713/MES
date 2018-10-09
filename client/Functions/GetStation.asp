<%
function getStation(isshowname,showstyle,station)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select NID,STATION_NAME from STATION order by STATION_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=station then
			output=output&" selected"
			end if
		output=output&">"&rsS("STATION_NAME")&"</option>"
		end select
	rsS.movenext
	wend
end if
getStation=output
rsS.close
set rsS=nothing
end function

function getStation_1(isshowname,showstyle,station)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""	
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select NID,STATION_NAME from STATION_New order by STATION_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=station then
			output=output&" selected"
			end if
		output=output&">"&rsS("STATION_NAME")&"</option>"
		end select
	rsS.movenext
	wend
end if
getStation_1=output
rsS.close
set rsS=nothing
end function

function getStation_New(isshowname,showstyle,station,where,order,sequency,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS3=server.CreateObject("adodb.recordset")
z=0
SQLS="Select S.NID,S.STATION_NUMBER,S.STATION_NAME,S.STATION_CHINESE_NAME,S.TRANSACTION_TYPE,F.FACTORY_NAME from STATION_New S inner join FACTORY F on S.FACTORY_ID=F.NID "& where & order
rsS3.open SQLS,conn,1,3
if not rsS3.eof then
	while not rsS3.eof
	z=z+1
	rsS3.movenext
	wend
end if
rsS3.close
SQLS="Select S.NID,STATION_NUMBER,S.STATION_NAME,S.STATION_CHINESE_NAME,S.TRANSACTION_TYPE,F.FACTORY_NAME from STATION_New S inner join FACTORY F on S.FACTORY_ID=F.NID  "&where & " and Is_delete<>'1' " &order
 
rsS3.open SQLS,conn,1,3
if not rsS3.eof then
	if sequency<>"" then
		aseq=split(sequency,",")
		for x=0 to ubound(aseq)
			while not rsS3.eof
				if aseq(x)=rsS3("NID") then
					select case showstyle
					case "OPTION"
						output=output&"<option value='"&rsS3("NID")&"'"
							if rsS3("NID")=station then
							output=output&"selected"
							end if
						output=output&">"&rsS3("STATION_NUMBER")&"&nbsp;"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&") - "&rsS3("FACTORY_NAME")&"</option>"
					case "CHECKBOX"
						output=output&"<input name='defectcode_station' type='checkbox' value='"&rsS3("NID")&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"&splitchar
						else
						output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
						end if
					case "ENGLISH_TEXT"
						if x<>ubound(aseq) then
						output=output&rsS3("STATION_NAME")&splitchar
						else
						output=output&rsS3("STATION_NAME")
						end if
					case "TEXT_LINK"
						if x<>ubound(aseq) then
						output=output&"<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&getStationTransaction(rsS3("TRANSACTION_TYPE"),rsS3("STATION_NAME"),rsS3("STATION_CHINESE_NAME"))&"</a>"&splitchar
						else
						output=output&"<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&getStationTransaction(rsS3("TRANSACTION_TYPE"),rsS3("STATION_NAME"),rsS3("STATION_CHINESE_NAME"))&"</a>"
						end if
					end select
				end if
			rsS3.movenext
			wend
		rsS3.movefirst
		next
	else
		y=1
		while not rsS3.eof
			select case showstyle
			case "OPTION"
				output=output&"<option value='"&rsS3("NID")&"'"
					if rsS3("NID")=station then
					output=output&"selected"
					end if
				output=output&">"&rsS3("STATION_NUMBER")&"&nbsp;" & rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&") - "&rsS3("FACTORY_NAME")&"</option>"
			case "CHECKBOX"
				output=output&"<input name='defectcode_station' type='checkbox' value='"&rsS3("NID")&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
			case "TEXT"
				if station<>"" then
					if rsS3("NID")=station then
					output=rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
					end if
				else
					if y<>z then
					output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"&splitchar
					else
					output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
					end if
				end if
			case "ENGLISH_TEXT"
				if station<>"" then
					if rsS3("NID")=station then
					output=rsS3("STATION_NAME")
					end if
				else
					if y<>z then
					output=output&rsS3("STATION_NAME")&splitchar
					else
					output=output&rsS3("STATION_NAME")
					end if
				end if
			case "TEXT_LINK"
				if station<>"" then
					if rsS3("NID")=station then
					output="<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")</a>"
					end if
				else
					if y<>z then
					output=output&"<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")</a>"&splitchar
					else
					output=output&"<a href='/Admin/Station/EditStation.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")</a>"
					end if
				end if
			end select
		rsS3.movenext
		y=y+1
		wend
	end if
end if
getStation_New=output
rsS3.close
set rsS3=nothing
end function




function getStation_New2(isshowname,showstyle,station,where,order,sequency,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS3=server.CreateObject("adodb.recordset")
z=0
SQLS="Select S.NID,S.STATION_NUMBER,S.STATION_NAME,S.STATION_CHINESE_NAME,S.TRANSACTION_TYPE,F.FACTORY_NAME from STATION_New S inner join FACTORY F on S.FACTORY_ID=F.NID where 1=1 "&where&order
 
rsS3.open SQLS,conn,1,3
if not rsS3.eof then
	while not rsS3.eof
	z=z+1
	rsS3.movenext
	wend
end if
rsS3.close
SQLS="Select S.NID,STATION_NUMBER,S.STATION_NAME,S.STATION_CHINESE_NAME,S.TRANSACTION_TYPE,F.FACTORY_NAME from STATION_New S inner join FACTORY F on S.FACTORY_ID=F.NID  where 1=1 "&where & " and Is_delete<>'1' " &order
 
rsS3.open SQLS,conn,1,3
if not rsS3.eof then
	if sequency<>"" then
		aseq=split(sequency,",")
		for x=0 to ubound(aseq)
			while not rsS3.eof
				if aseq(x)=rsS3("NID") then
					select case showstyle
					case "OPTION"
						output=output&"<option value='"&rsS3("NID")&"'"
							if rsS3("NID")=station then
							output=output&"selected"
							end if
						output=output&">"&rsS3("STATION_NUMBER")&"&nbsp;"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&") - "&rsS3("FACTORY_NAME")&"</option>"
					case "CHECKBOX"
						output=output&"<input name='defectcode_station' type='checkbox' value='"&rsS3("NID")&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"&splitchar
						else
						output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
						end if
					case "ENGLISH_TEXT"
						if x<>ubound(aseq) then
						output=output&rsS3("STATION_NAME")&splitchar
						else
						output=output&rsS3("STATION_NAME")
						end if
					case "TEXT_LINK"
						if x<>ubound(aseq) then
						output=output&"<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&getStationTransaction(rsS3("TRANSACTION_TYPE"),rsS3("STATION_NAME"),rsS3("STATION_CHINESE_NAME"))&"</a>"&splitchar
						else
						output=output&"<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&getStationTransaction(rsS3("TRANSACTION_TYPE"),rsS3("STATION_NAME"),rsS3("STATION_CHINESE_NAME"))&"</a>"
						end if
					end select
				end if
			rsS3.movenext
			wend
		rsS3.movefirst
		next
	else
		y=1
		while not rsS3.eof
			select case showstyle
			case "OPTION"
				output=output&"<option value='"&rsS3("NID")&"'"
					if rsS3("NID")=station then
					output=output&"selected"
					end if
				output=output&">"&rsS3("STATION_NUMBER")&"&nbsp;" & rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&") - "&rsS3("FACTORY_NAME")&"</option>"
			case "CHECKBOX"
				output=output&"<input name='defectcode_station' type='checkbox' value='"&rsS3("NID")&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
			case "TEXT"
				if station<>"" then
					if rsS3("NID")=station then
					output=rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
					end if
				else
					if y<>z then
					output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"&splitchar
					else
					output=output&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")"
					end if
				end if
			case "ENGLISH_TEXT"
				if station<>"" then
					if rsS3("NID")=station then
					output=rsS3("STATION_NAME")
					end if
				else
					if y<>z then
					output=output&rsS3("STATION_NAME")&splitchar
					else
					output=output&rsS3("STATION_NAME")
					end if
				end if
			case "TEXT_LINK"
				if station<>"" then
					if rsS3("NID")=station then
					output="<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")</a>"
					end if
				else
					if y<>z then
					output=output&"<a href='/Admin/Station/EditStation_New.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")</a>"&splitchar
					else
					output=output&"<a href='/Admin/Station/EditStation.asp?id="&rsS3("NID")&"&path="&path&"&query="&query&"'>"&rsS3("STATION_NAME")&"&nbsp;("&rsS3("STATION_CHINESE_NAME")&")</a>"
					end if
				end if
			end select
		rsS3.movenext
		y=y+1
		wend
	end if
end if
getStation_New2=output
rsS3.close
set rsS3=nothing
end function
%>