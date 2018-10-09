<%
function getFamily(showstyle,Family,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select S.NID,S.FAMILY_NAME from FAMILY S inner join FACTORY F on S.FACTORY_ID=F.NID "&where& " ORDER BY FAMILY_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=Family then
			output=output&"selected"
			end if
		output=output&">"&rsS("FAMILY_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='series' type='radio' class='t-c-greenCopy' value='"&rsS("NID")&"'>"&rsS("FAMILY_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsS("FAMILY_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getFamily=output
rsS.close
set rsS=nothing
end function

function getSeries(showstyle,Series,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select S.NID,S.SERIES_NAME from SERIES_NEW S inner join FACTORY F on S.FACTORY_ID=F.NID "&where& " ORDER BY SERIES_NAME"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=Series then
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


function getSubSeries(showstyle,SubSeries,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select S.* from SUBSERIES S inner join FACTORY F on S.FACTORY_ID=F.NID "&where& " ORDER BY SUBSERIES_NAME"

rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("NID")&"'"
			if rsS("NID")=SubSeries then
			output=output&"selected"
			end if
		output=output&">"&rsS("SUBSERIES_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='series' type='radio' class='t-c-greenCopy' value='"&rsS("NID")&"'>"&rsS("SUBSERIES_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsS("SUBSERIES_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getSubSeries=output
rsS.close
set rsS=nothing
end function


function getModel(showstyle,ItemName,where,order,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select distinct item_name from PRODUCT_MODEL "&where&" order by item_name"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("item_name")&"'"
			if rsS("item_name")=ItemName then
			output=output&"selected"
			end if
		output=output&">"&rsS("item_name")&"</option>"
		case "RADIO"
		output=output&"<input name='series' type='radio' class='t-c-greenCopy' value='"&rsS("ItemName")&"'>"&rsS("ItemName")&"&nbsp;"
		case "TEXT"
			output=output&rsS("ItemName")&splitchar
		end select
	rsS.movenext
	wend
end if
getModel=output
rsS.close
set rsS=nothing
end function




function getIncludeModel(SubSeriesID)
 
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select distinct item_name from PRODUCT_MODEL WHERE SERIES_ID='"& SubSeriesID &"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		output=output& rsS("item_name").value &","
		rsS.movenext
	wend
end if
getIncludeModel=output
rsS.close
set rsS=nothing
end function



function get_Start_DateTime(wwStart)
		get_Start_DateTime=cint(wwStart-1)*7+cdate("2010-01-02 23:59:59")
end function 

function get_End_DateTime(wwEnd)
		get_End_DateTime=(cint(wwEnd)-1)*7+cdate("2010-01-09 23:59:59")
end function 


function get_Start_DateTime_New(wwStart,mYear)
	 if(mYear="2010") then
		get_Start_DateTime_New=cint(wwStart-1)*7+cdate("2010-01-02 23:59:59")
	 end if
	 if(mYear="2011") then
		get_Start_DateTime_New=cint(wwStart-1)*7+cdate("2011-01-01 23:59:59")
	 end if

	  if(mYear="2012") then
		get_Start_DateTime_New=cint(wwStart-1)*7+cdate("2011-12-31 23:59:59")
	 end if
	 
	  if(mYear="2013") then
		get_Start_DateTime_New=cint(wwStart-1)*7+cdate("2012-12-30 23:59:59")
	 end if
	 
	 
end function 

function get_End_DateTime_New(wwEnd,mYear)
	  if(mYear="2010") then
		get_End_DateTime_New=(cint(wwEnd)-1)*7+cdate("2010-01-09 23:59:59")
	 end if
	 if(mYear="2011") then
		get_End_DateTime_New=(cint(wwEnd)-1)*7+cdate("2011-01-08 23:59:59")
	 end if

	 if(mYear="2012") then
		get_End_DateTime_New=(cint(wwEnd)-1)*7+cdate("2012-01-07 23:59:59")
	 end if
	 
	  if(mYear="2013") then
		get_End_DateTime_New=cint(wwEnd-1)*7+cdate("2012-12-30 23:59:59")
	 end if
	 
end function 



%>
