<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Station/StationCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
'response.Write("s:"&request.Form("id"))
'response.End()
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
id=request.Form("id")
stationnumber=trim(request.Form("stationnumber"))
stationname=trim(request.Form("stationname"))
group_nids=replace(request.Form("toitem3")," ","")
'Tolerance=trim(request.form("Tolerance"))
'ToleranceUnit=trim(request.form("ToleranceUnit"))
Tolerance=1
ToleranceUnit="H"

SQL="select * from STATION_New where NID='"&id&"'"

rs.open SQL,conn,1,3
if not rs.eof then
	rs("STATION_NUMBER")=stationnumber
	rs("STATION_NAME")=stationname
	rs("STATION_CHINESE_NAME")=trim(request.Form("stationchinesename"))
	if request.Form("section")<>"" then
		rs("SECTION_ID")=request.Form("section")
	end if
	if request.Form("factory")<>"" then
		rs("FACTORY_ID")=request.Form("factory")
	end if
	if request.Form("actionscount")<>"" then
		rs("ACTIONS_COUNT")=request.Form("actionscount")
	end if
	rs("INITAIL_QUANTITY_TYPE")=request.Form("quantity_type")
	
	rs("WIP_REPORT_COLUMN")=request.Form("WIP")
	rs("WIP_INCLUDED_STATIONS")=replace(request.Form("toitem2")," ","")
	rs("OUTPUT_REPORT_COLUMN")=request.Form("Output")
	if request.Form("transaction")<>"" then
		rs("TRANSACTION_TYPE")=cstr(request.Form("transaction"))
	else
		rs("TRANSACTION_TYPE")="0"
	end if
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
		rs("IS_DELETE")=0
		rs("LASTUPDATE_PERSON")=session("code")
		rs("LASTUPDATE_TIME")=date()
		
		rs("Tolerance")=Tolerance
		rs("ToleranceUnit")=ToleranceUnit
		rs.update
		rs.close
		word="Successfully edit Station."
		action="location.href='"&beforepath&"'"
		
		
	'add by jack zhang 2011-1-17 update releated action 
	SQL="UPDATE STATION SET STATION_NAME='"+stationname+"',"
	SQL=SQL+"STATION_NUMBER='"+stationnumber+"',"
	SQL=SQL+"INITAIL_QUANTITY_TYPE='"+request.Form("quantity_type")+"',"
	SQL=SQL+"STATION_CHINESE_NAME='"+trim(request.Form("stationchinesename"))+"',"
	if request.Form("factory")<>"" then
		SQL=SQL+"FACTORY_ID='"+trim(request.Form("factory"))+"',"
	ELSE
		SQL=SQL+"FACTORY_ID=NULL,"
	END IF 
	
	if request.Form("section")<>"" then
		SQL=SQL+"SECTION_ID='"+trim(request.Form("section"))+"',"
	end if
	if request.Form("transaction")<>"" then
		SQL=SQL+"TRANSACTION_TYPE='"+cstr(request.Form("transaction"))+"'"
		 
	else
		SQL=SQL+"TRANSACTION_TYPE='0'"
	end if
	SQL=SQL+" WHERE MOTHER_STATION_ID='"+id+"'"
	set rs0=server.createobject("adodb.recordset")
	rs0.open sql,conn,1,3
	'END ADD
else
	word="Station of "&stationname&" has not existed, please input again."
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
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->