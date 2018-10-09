<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/SubJobs/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetStationsTransaction.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
query=replace(query,"*","&")
jobnumber=request.QueryString("jobnumber")
sheetnumber=request.QueryString("sheetnumber")
jobtype=request.QueryString("jobtype")
reason=request.QueryString("reason")
SQL="select STATUS,STATIONS_INDEX,CURRENT_STATION_ID from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
current_station_id=rs("CURRENT_STATION_ID")
stations_index=rs("STATIONS_INDEX")
rs.close
SQL="select CONTROL_TYPE,CONTROL_STATION,CONTROL_REASON,CONTROL_PERSON,CONTROL_TIME from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
rs("CONTROL_TYPE")=rs("CONTROL_TYPE")&"5"&","
rs("CONTROL_STATION")=rs("CONTROL_STATION")&current_station_id&","
rs("CONTROL_REASON")=rs("CONTROL_REASON")&reason&","
rs("CONTROL_PERSON")=rs("CONTROL_PERSON")&session("code")&","
rs("CONTROL_TIME")=rs("CONTROL_TIME")&now()&","
rs.update
rs.close

station_name=getStation(true,"TEXT",stations_index,"","",stations_index,",")
station_transaction=getStationsTransaction(stations_index)
astation_index=split(stations_index,",")
astation_name=split(station_name,",")
astation_transaction=split(station_transaction,",")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
</head>

<body>
<form id="form1" name="form1" method="post" action="/Job/SubJobs/RepeatJob1.asp">
<table width="376" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#003366" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="3" class="t-c-greenCopy">Please select stations need to repeat </td>
    </tr>
  <tr>
    <td height="20" colspan="3" class="t-t-Borrow">Job Number:<%= jobnumber %>&nbsp;</td>
    </tr>
<%for i=0 to ubound(astation_index)%>
    <tr>
      <td width="74"><div align="center"><%= i+1%>&nbsp;</div></td>
    <td width="50" height="20">
      <div align="center">
        <input name="station_id" type="checkbox" id="station_id" value="<%=astation_index(i)%>" <%if astation_transaction(i)="0" then%>disabled<%end if%> />    
      </div></td>
    <td width="244" height="20"><%= astation_name(i) %>&nbsp;</td>
  </tr>
<%next%>
  <tr>
    <td height="20" colspan="3"><div align="center">
      <input name="jobnumber" type="hidden" id="jobnumber" value="<%=jobnumber%>">
      <input name="sheetnumber" type="hidden" id="sheetnumber" value="<%=sheetnumber%>">
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
<input type="submit" name="Submit" value="Submit">
&nbsp;
      <input type="reset" name="Submit2" value="Reset">
    </div></td>
    </tr>
</table>
</form>
</body>
</html><!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->