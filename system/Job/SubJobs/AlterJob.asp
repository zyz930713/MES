<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Functions/FormatFunctions.asp" -->
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
SQL="select STATUS,STATIONS_INDEX,CURRENT_STATION_ID,FINISHED_STATIONS_ID from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
current_station_id=rs("CURRENT_STATION_ID")
stations_index=rs("STATIONS_INDEX")
finished_stations_id=rs("FINISHED_STATIONS_ID")
rs.close
SQL="select CONTROL_TYPE,CONTROL_STATION,CONTROL_REASON,CONTROL_PERSON,CONTROL_TIME from JOB where JOB_NUMBER='"&jobnumber&"' and SHEET_NUMBER='"&sheetnumber&"' and JOB_TYPE='"&jobtype&"'"
rs.open SQL,conn,1,3
rs("CONTROL_TYPE")=rs("CONTROL_TYPE")&"6"&","
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
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
</head>

<body>
<form id="form1" name="form1" method="post" action="/Job/SubJobs/RepeatJob1.asp">
<table width="376" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#003366" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" class="t-c-greenCopy">Please Alter Job </td>
    </tr>
  <tr>
    <td height="20" class="t-t-Borrow">Job Number:<%= jobnumber %>-<%=replace(jobtype,"N","")&repeatstring(sheetnumber,"0",3)%>&nbsp;</td>
    </tr>

    <tr>
      <td height="20"> Stations' Sequence </td>
    </tr>
    <tr>
      <td width="244" height="20">
        <table  border="1" align="left" cellpadding="0" cellspacing="0" bordercolorlight="#73A2EE" bordercolordark="#FFFFFF">
          <tr>
            <td>
              <%for i=0 to ubound(astation_index)
		  if instr(finished_stations_id,astation_index(i))=0 then
		  	current_flag=i
			exit for
		  else%>
              <%= astation_name(i) %>
              <%
		  end if
		  next%></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td rowspan="4"><select name="stationindex" size="10" id="stationindex">
              <%for i=current_flag to ubound(astation_index)%>
              <option value="<%=astation_index(i)%>" <%if instr(finished_stations_id,astation_index(i))<>0 then%>disabled<%end if%>><%= astation_name(i) %></option>
              <%next%>
            </select></td>
          <td><div align="center"> <img src="/Images/Button_Up.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_up(document.form1.stationindex)"> </div></td>
        </tr>
        
        <tr>
          <td><div align="center"> <img src="/Images/Button_Down.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_down(document.form1.stationindex)"> </div></td>
        </tr>
        
        <tr>
          <td><div align="center"> <img src="/Images/Button_To_Top.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_top(document.form1.stationindex)"> </div></td>
        </tr>
        
        <tr>
          <td><div align="center"> <img src="/Images/Button_To_Bottom.gif" width="73" height="18" align="absmiddle" style="cursor:hand" onClick="item_bottom(document.form1.stationindex)"> </div></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="20"><div align="center">
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