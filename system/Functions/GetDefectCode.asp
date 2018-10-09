<%
function getDefectCode(showtype,defectcode,where,order,splitchar)
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select D.NID,D.DEFECT_CODE,D.DEFECT_NAME,D.DEFECT_CHINESE_NAME,F.FACTORY_NAME from DEFECTCODE D inner join FACTORY F on D.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
while not rsP.eof
	select case showtype
	case "OPTION"
	output=output&"<option value='"&rsP("NID")&"'>"&rsP("DEFECT_CODE")&" - "&rsP("DEFECT_NAME")&"("&rsP("DEFECT_CHINESE_NAME")&") - "&rsP("FACTORY_NAME")&"</option>"
	case "TEXT"
	output=output&rsP("DEFECT_NAME")&"("&rsP("DEFECT_CHINESE_NAME")&")"&splitchar
	end select
rsP.movenext
wend
getDefectCode=output
rsP.close
set rsP=nothing
end function

function getDefectCode_New(showtype,defectcode,where,order,splitchar)
output=""
set rsP=server.CreateObject("adodb.recordset")
order=" order by DEFECT_CODE "
SQLP="Select D.NID,D.DEFECT_CODE,D.DEFECT_NAME,D.DEFECT_CHINESE_NAME,F.FACTORY_NAME from DEFECTCODE_New D inner join FACTORY F on D.FACTORY_ID=F.NID where IS_DELETE<>1 "&where&order

rsP.open SQLP,conn,1,3
while not rsP.eof
	select case showtype
	case "OPTION"
	addString=""
	if rsp("NID")=defectcode then
		addString="selected"
	end if
	output=output&"<option value='"&rsP("NID")&"' "&addString&">"&rsP("DEFECT_CODE")&" - "&rsP("DEFECT_NAME")&"("&rsP("DEFECT_CHINESE_NAME")&") - "&rsP("FACTORY_NAME")&"</option>"
	case "TEXT"
	output=output&rsP("DEFECT_NAME")&"("&rsP("DEFECT_CHINESE_NAME")&")"&splitchar
	end select
rsP.movenext
wend
getDefectCode_New=output
rsP.close
set rsP=nothing
end function
%>