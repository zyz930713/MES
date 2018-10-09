<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/Event_Transaction.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
seriesgroupname=trim(request.Form("seriesgroupname"))
series_nids=replace(request.Form("toitem")," ","")
SQL="select * from SERIES_GROUP where SERIES_GROUP_NAME='"&seriesgroupname&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="SG"&NID_SEQ("SERIES_GROUP")
rs("SERIES_GROUP_NAME")=seriesgroupname
rs("FACTORY_ID")=request.Form("factory")
rs("SECTION_ID")=request.Form("section")
'rs("LINE_ID")=request.Form("line")
rs("INCLUDED_SERIES")=series_nids
	if series_nids<>"" then
	included_system_items=""
		set rsS=server.CreateObject("adodb.recordset")
		SQLS="select INCLUDED_SYSTEM_ITEMS from SERIES where NID in ('"&replace(series_nids,",","','")&"') order by SERIES_NAME"
		rsS.open SQLS,conn,1,3
		if not rsS.eof then
			while not rsS.eof
			included_system_items=included_system_items&rsS("INCLUDED_SYSTEM_ITEMS")&","
			rsS.movenext
			wend
		end if
		rsS.close
		set rsS=nothing
	end if
	if included_system_items<>"" then
	included_system_items=left(included_system_items,len(included_system_items)-1)
	end if
rs("INCLUDED_SYSTEM_ITEMS")=included_system_items
if request.Form("overall_exception")="1" then
rs("OVERALL_EXCEPTION")="1"
else
rs("OVERALL_EXCEPTION")="0"
end if
lead_time=timeconvert(request.Form("lead_time"),request.Form("lead_time_unit"))
rs("LEAD_TIME")=lead_time
wip_time=timeconvert(request.Form("wip_time"),request.Form("wip_time_unit"))
response.Write(wip_time)
rs("WIP_TIME")=wip_time
rs("TARGET_FIRSTYIELD")=request.Form("first_yield")
rs("TARGET_INTERNALYIELD")=csng(request.Form("internal_yield"))/100
rs("TARGET_INSPECTYIELD")=csng(request.Form("inspect_yield"))/100
rs("TARGET_YIELD")=request.Form("final_yield")
doEvent "SERIES_GROUP_UPDATE",jobnumber,series_name,seriesgroupname
rs.update
rs.close
word="Successfully save a New Series Group."
action="location.href='"&beforepath&"'"
else
word="Series Group of "&seriesgroupname&" has existed, please input again."
action="history.back()"
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->