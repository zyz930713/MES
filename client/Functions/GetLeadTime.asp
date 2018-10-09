<%
function getLeadTime(jobnumber,part_number_tag,factory_id)
set rsS=server.CreateObject("adodb.recordset")
SQLS="Select SERIES_GROUP_NAME,LEAD_TIME from SERIES_GROUP where INCLUDED_SYSTEM_ITEMS like '%"&part_number_tag&"%'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	if rsS.recordcount=1 then
		getLeadTime=csng(rsS("LEAD_TIME"))
'		if getLeadTime=0 then
'		sendEmail "BarcodeSystem@knowles.com","Judy.Wang@knowles.com","Wrong Settings of Series Group for "&rsS("SERIES_GROUP_NAME"),"Hi, administrator<p>Lead Time is not set in series group ["&rsS("SERIES_GROUP_NAME")&"] when barcode scan job of "&jobnumber&", please update.<p>Barcode System"
'		end if
	else
		if instr(jobnumber,"E")<=0 then
			getLeadTime=0
			while not rsS.eof 
			series_groups=series_groups&rsS("SERIES_GROUP_NAME")&","
			rsS.movenext
			wend
			series_groups=left(series_groups,len(series_groups)-1)
'			sendEmail "BarcodeSystem@knowles.com","Judy.Wang@knowles.com","Wrong Settings of Series Group for "&series_groups,"Hi, administrator<p>"&part_number_tag&" is categorized into multiple series groups ["&series_groups&"] when barcode scan job of "&jobnumber&", please update.<p>Barcode System"
		else
			getLeadTime=0
		end if
	end if
else
	getLeadTime=0
'	if factory_id="FA00000002" and instr(jobnumber,"E")<=0 then
'	sendEmail "BarcodeSystem@knowles.com","Judy.Wang@knowles.com","Wrong Settings of Series Group","Hi, administrator<p>"&part_number_tag&" is not categorized into any series group when barcode scan job of "&jobnumber&", please update.<p>Barcode System"
'	end if
end if
rsS.close
set rsS=nothing
end function
%>