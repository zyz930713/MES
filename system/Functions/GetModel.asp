<%
function getModel(showstyle,model,where,order,splitchar,idcount)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
model=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select 1,M.ITEM_ID,M.ITEM_NAME,F.FACTORY_NAME from PRODUCT_MODEL M left join FACTORY F on M.FACTORY_ID=F.NID where M.ITEM_ID is not null "&where&order
'response.Write(SQLP)
session("aerror")=SQLP
'response.End()
rsP.open SQLP,conn,1,3
if not rsP.eof then
idcount=rsP.recordcount
	while not rsP.eof
		select case showstyle
		case "OPTION"
		'if instr(model,rsP("ITEM_NAME"))<=0 then
			output=output&"<option value='"&rsP("ITEM_NAME")&"'"
				if rsP("ITEM_NAME")=model then
				output=output&"selected"
				end if	
			output=output&">"&rsP("ITEM_NAME")&"&nbsp;("&rsP("FACTORY_NAME")&")</option>"
			model=model&rsP("ITEM_NAME")&","
		'end if
		case "RADIO"
		output=output&"<input name='series' type='radio' class='t-c-greenCopy' value='"&rsP("ITEM_NAME")&"'>"&rsP("ITEM_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsP("ITEM_NAME")&splitchar
		end select
	rsP.movenext
	wend
end if
getModel=output
rsP.close
set rsP=nothing
end function
%>