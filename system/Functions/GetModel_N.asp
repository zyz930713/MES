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
SQLP="Select PM.ITEM_ID,PM.ITEM_NAME,PM.ITEM_TYPE,PM.ITEM_STATUS,F.FACTORY_NAME from PRODUCT_MODEL PM left join FACTORY F on PM.FACTORY_ID=F.NID where PM.ITEM_NAME is not null "&where&order
rsP.open SQLP,conn,1,3
if not rsP.eof then
idcount=rsP.recordcount
	while not rsP.eof
	organization_id=""
		select case showstyle
		case "OPTION"
		if instr(model,rsP("ITEM_NAME"))<=0 then
			output=output&"<option value='"&rsP("ITEM_ID")&"'"
				if rsP("ITEM_NAME")=model then
				output=output&"selected"
				end if
			output=output&">"&rsP("ITEM_NAME")&"&nbsp;("&rsP("FACTORY_NAME")&")</option>"
			model=model&rsP("ITEM_NAME")&","
		end if
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