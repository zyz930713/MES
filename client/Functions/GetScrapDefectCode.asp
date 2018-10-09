<%
function getScrapDefectCode(factory,factory_name)
set rsS=server.CreateObject("adodb.recordset")
select case factory_name
case "KE"
SQLD="Select DC.NID,DC.DEFECT_CODE,DC.DEFECT_NAME,DC.DEFECT_CHINESE_NAME from DEFECTCODE DC inner join FACTORY F on DC.FACTORY_ID=F.NID where DC.STATION_ID='ST00000049' and F.NID='"&factory&"' order by DC.NID"
case "EMC"
SQLD="Select DC.NID,DC.DEFECT_CODE,DC.DEFECT_NAME,DC.DEFECT_CHINESE_NAME from DEFECTCODE DC inner join FACTORY F on DC.FACTORY_ID=F.NID where (DC.STATION_ID='ST00000034' or DC.STATION_ID='ST00000033')and F.NID='"&factory&"' order by DC.NID"
case "VAM"
SQLD="Select DC.NID,DC.DEFECT_CODE,DC.DEFECT_NAME,DC.DEFECT_CHINESE_NAME from DEFECTCODE DC inner join FACTORY F on DC.FACTORY_ID=F.NID where (DC.STATION_ID='ST00000326' or DC.STATION_ID='ST00000074' or DC.STATION_ID='ST00000329') and F.NID='"&factory&"' order by DC.NID"
end select
rsS.open SQLD,conn,1,3
while not rsS.eof 
	optiontext=optiontext&"<option value='"&rsS("DEFECT_CHINESE_NAME")&"'>"&left(rsS("DEFECT_CHINESE_NAME"),10)&"</option>"
rsS.movenext
wend
getScrapDefectCode=optiontext
rsS.close
set rsS=nothing
end function
%>