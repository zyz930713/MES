<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Part/PartCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/GetStationsTransaction.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
partnumber=trim(request.Form("partnumber"))
part_rule=trim(request.Form("part_rule"))

SQL="select * from ROUTING where PART_NUMBER='"&partnumber&"'"
rs.open SQL,conn,1,3
if not rs.eof then
	word="Part of "&partnumber&" has existed, please input again."
	action="history.back()"
else
	rs.addnew
	rs("NID")="RT"&NID_SEQ("ROUTING")
	rs("JOB_RULE")=trim(request.Form("job_rule"))
	rs("PART_NUMBER")=partnumber
	if right(part_rule,1)="," then
		this_part_rule=left(part_rule,len(part_rule)-1)
	else
		this_part_rule=part_rule
	end if
	rs("PART_RULE")=this_part_rule
	rs("PART_TYPE")=request.Form("partType")
	rs("FACTORY_ID")=request.Form("factory")
	rs("SECTION_ID")=request.Form("section")
	rs("STATIONS_COUNT")=request.Form("stationscount")
	lines_index=replace(request.Form("toitem1")," ","")
	rs("LINES_INDEX")=lines_index
	stations_index=replace(request.Form("toitem2")," ","")
	rs("STATIONS_INDEX")=stations_index
	rs("STATIONS_ROUTINE")=request.Form("stationroutine")
	rs("STATIONS_TRANSACTION")=getStationsTransaction(stations_index)
	rs("MAX_INTERVAL")=replace(request.Form("maxinterval")," ","")
	
	if request.Form("initial_quantity")<>"" then
		rs("INITIAL_QUANTITY")=trim(request.Form("initial_quantity"))
	end if
	rs("TARGET_YIELD")=request.Form("yield")
	rs("MEET_PRIORITY")=request.Form("priority")
	rs("ROUTINE_TYPE")=null
	rs("ROUTINE_PURPOSE")=null
	rs.update
	rs.close
	word="Successfully save a New Part."
	action="location.href='"&beforepath&"'"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->