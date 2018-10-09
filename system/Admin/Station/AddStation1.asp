<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
stationnumber=trim(request.Form("stationnumber"))
stationname=trim(request.Form("stationname"))
group_nids=replace(request.Form("toitem3")," ","")
SQL="select * from STATION where STATION_NUMBER='"&stationnumber&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("NID")="ST"&NID_SEQ("STATION")
rs("STATION_NUMBER")=stationnumber
rs("STATION_NAME")=stationname
rs("STATION_CHINESE_NAME")=request.Form("stationchinesename")
if request.Form("section")<>"" then
rs("SECTION_ID")=request.Form("section")
end if
if request.Form("factory")<>"" then
rs("FACTORY_ID")=request.Form("factory")
end if
if request.Form("actionscount")<>"" then
rs("ACTIONS_COUNT")=request.Form("actionscount")
end if
rs("ACTIONS_INDEX")=replace(request.Form("toitem")," ","")
rs("INITAIL_QUANTITY_TYPE")=request.Form("quantity_type")
included_defecodes=""
if group_nids<>"" then
	set rsS=server.CreateObject("adodb.recordset")
	SQLS="select MEMBERS_ID from DEFECTCODE_GROUP where NID in ('"&replace(group_nids,",","','")&"') order by GROUP_NAME"
	rsS.open SQLS,conn,1,3
	if not rsS.eof then
		while not rsS.eof
			if rsS("MEMBERS_ID")<>"" then
				a_members_id=split(rsS("MEMBERS_ID"),",")
				for i=0 to ubound(a_members_id)
					if instr(included_defecodes,a_members_id(i))<=0 then
					included_defecodes=included_defecodes&a_members_id(i)&","
					end if
				next
			end if
		rsS.movenext
		wend
	end if
	rsS.close
	set rsS=nothing
	if included_defecodes<>"" then
	included_defecodes=left(included_defecodes,len(included_defecodes)-1)
	end if
end if
rs("DEFECTCODE_GROUPS")=group_nids
rs("DEFECTCODES_INDEX")=included_defecodes
rs("WIP_REPORT_COLUMN")=request.Form("WIP")
rs("WIP_INCLUDED_STATIONS")=replace(request.Form("toitem2")," ","")
rs("STATION_ENTER_DEFECTCODE")=replace(request.Form("toitem1")," ","")
rs("OUTPUT_REPORT_COLUMN")=request.Form("Output")
if request.Form("transaction")<>"" then
rs("TRANSACTION_TYPE")=cstr(request.Form("transaction"))
else
rs("TRANSACTION_TYPE")="0"
end if
rs("TRANSACTION_CHANGE")=request.Form("ischange")
if request.Form("WIP_SEQ")<>"" then
rs("WIP_SEQUENCY")=request.Form("WIP_SEQ")
else
rs("WIP_SEQUENCY")=null
end if
if request.Form("Output_SEQ")<>"" then
rs("OUTPUT_SEQUENCY")=request.Form("Output_SEQ")
else
rs("OUTPUT_SEQUENCY")=null
end if
rs.update
word="Successfully save New Sation."
action="location.href='"&beforepath&"'"
else
word="Sation of "&stationname&" has existed, please input again."
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