<%
function getMaterialCategory(showstyle,CategoryID,where,order,splitchar)
output=""
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select  *  from MATERIAL_CATEGORY "&where&order
rsS.open SQLS,conn,1,3
if not rsS.eof then
	while not rsS.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsS("CATEGORY_ID")&"'"
			if rsS("CATEGORY_ID")=CategoryID then
			output=output&"selected"
			end if
		output=output&">"&rsS("CATEGORY_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='CATEGORY' type='radio' class='t-c-greenCopy' value='"&rsS("CATEGORY_ID")&"'>"&rsS("CATEGORY_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsS("CATEGORY_NAME")&splitchar
		end select
	rsS.movenext
	wend
end if
getMaterialCategory=output
rsS.close
set rsS=nothing
end function
%>