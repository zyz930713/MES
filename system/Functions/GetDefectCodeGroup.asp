<%
function getDefectCodeGroup(showtype,defectcode,where,order,splitchar)
output=""
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select D.NID,D.GROUP_NAME,D.GROUP_CHINESE_NAME,F.FACTORY_NAME from DEFECTCODE_GROUP D inner join FACTORY F on D.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
while not rsP.eof
	select case showtype
	case "OPTION"
	output=output&"<option value='"&rsP("NID")&"'>"&rsP("GROUP_NAME")&"("&rsP("GROUP_CHINESE_NAME")&") - "&rsP("FACTORY_NAME")&"</option>"
	case "TEXT"
	output=output&rsP("GROUP_NAME")&"("&rsP("GROUP_CHINESE_NAME")&")"&splitchar
	end select
rsP.movenext
wend
getDefectCodeGroup=output
rsP.close
set rsP=nothing
end function
%>