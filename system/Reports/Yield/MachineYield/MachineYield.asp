<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetMachine.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<!--#include virtual="/Functions/GetMachineStationDefectCode.asp" -->
<!--#include virtual="/Functions/GetJobTotalDefectCodeQuantity.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsJ=server.CreateObject("adodb.recordset")
set rsU=server.CreateObject("adodb.recordset")
machine=request("machine")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc, J.SHEET_NUMBER"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if machine<>"" and machine<>"all" then
mwhere=" and NID='"&machine&"'"
end if
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER='"&jobnumber&"'"
end if
if partnumber<>"" then
where=where&" and J.PART_NUMBER_TAG like '%"&partnumber&"%'"
end if
if fromdate<>"" then
where=where&" and START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
else
fromdate=date()
where=where&" and START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and START_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&machine="&machine&"&jobnumber="&jobnumber&"&partnumber="&partnumber&"&fromdate="&fromdate&"&todate="&todate&"&ordername="&ordername&"&ordertype="&ordertype
pagename="/Reports/Yield/MachineYield/MachineYield.asp"
FactoryRight ""
SQL="select MACHINE_NUMBER,MACHINE_NAME,STATIONS_USED from MACHINE where STATIONS_USED is not null "&mwhere&factorywhereoutsideand&" ORDER BY MACHINE_NUMBER"
session("SQL")=SQL
session("where")=where
session("order")=order
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<form name="form1" method="post" action="/Reports/Yield/MachineYield/MachineYield.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span>Search Job </span></td>
  </tr>
  <tr>
    <td>Machine</td>
    <td><select name="machine" id="machine">
      <option value="">All</option>
	  <%FactoryRight ""%>
      <%=getMachine(true,"OPTION",machine,factorywhereoutside," order by MACHINE_NUMBER","")%>
    </select></td>
    <td height="20"><span class="style1">Job Number</span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td>Part Number </td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td>Station Start Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy">Browse Machine Yield </td>
</tr>
<tr>
  <td height="20" colspan="10" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
        <% =session("User") %></td>
      <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('MachineYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="10"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<%
if not rs.eof then
	i=1
	p=1
	rs.absolutepage=session("strpagenum")
	while not rs.eof and i<=rs.pagesize
		SQLJ="select 1,count(J.JOB_NUMBER) as idcount from JOB J inner join JOB_ACTIONS JA on J.JOB_NUMBER=JA.JOB_NUMBER and J.SHEET_NUMBER=JA.SHEET_NUMBER where JA.ACTION_VALUE='"&rs("MACHINE_NUMBER")&"' "&where
		rsJ.open SQLJ,conn,1,3
		q=rsJ("idcount")
		rsJ.close
%>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><%=i%></div></td>
  <td class="t-t-Borrow"><div align="center"><%=rs("MACHINE_NUMBER")%></div></td>
  <td height="20" class="t-t-Borrow"><div align="center">Job Number
    </div></td>
  <td class="t-t-Borrow"><div align="center">Part Number
    </div></td>
  <%astations=split(rs("STATIONS_USED"),",")
  for s=0 to ubound(astations)%>
  <td class="t-t-Borrow"><div align="center"><%=getStation(true,"TEXT",astations(s),"","","","")%>&nbsp;</div></td>
  <td class="t-t-Borrow"><div align="center">Start Time
    </div></td>
  <td class="t-t-Borrow"><div align="center"> Yield </div></td>
  <%next%>
  </tr>
<%		if q<>0 then
		SQLJ="select 1,J.JOB_NUMBER,J.SHEET_NUMBER,J.START_TIME,J.PART_NUMBER_TAG from JOB J inner join JOB_ACTIONS JA on J.JOB_NUMBER=JA.JOB_NUMBER and J.SHEET_NUMBER=JA.SHEET_NUMBER where JA.ACTION_VALUE='"&rs("MACHINE_NUMBER")&"' "&where&order
		rsJ.open SQLJ,conn,1,3
		if not rsJ.eof then
		while not rsJ.eof 
%>
<tr>
  	<%if p=1 then%><td height="20" rowspan="<%=q%>"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td><%end if%>
    <%if p=1 then%><td rowspan="<%=q%>"><div align="center"><%=rs("MACHINE_NUMBER")%>&nbsp;</div></td><%end if%>
    <td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rsJ("JOB_NUMBER")%>&sheetnumber=<%=rsJ("SHEET_NUMBER")%>" target="_blank"><%= rsJ("JOB_NUMBER") %>-<%=repeatstring(rsJ("SHEET_NUMBER"),"0",3)%></a></div></td>
    <td height="20"><div align="center"><%= rsJ("PART_NUMBER_TAG") %>&nbsp;</div></td>
	<%
			for s=0 to ubound(astations)
			stationstartquantity=0
			stationtotaldefectcodequantity=0
			station_assembly_yield=0
			station_start_time=""
			txtDF="&nbsp;"
			SQLU="select START_TIME,STATION_START_QUANTITY,STATION_DEFECTCODE_QUANTITY,STATION_ASSEMBLY_YIELD from JOB_STATIONS where JOB_NUMBER='"&rsJ("JOB_NUMBER")&"' and SHEET_NUMBER='"&rsJ("SHEET_NUMBER")&"' and STATION_ID='"&astations(s)&"'"
			rsU.open SQLU,conn,1,3
			if not rsU.eof then
			station_start_time=rsU("START_TIME")
				if cint(rsU("STATION_ASSEMBLY_YIELD"))=0 then
				stationstartquantity=getJobActionValue(rsJ("JOB_NUMBER"),rsJ("SHEET_NUMBER"),astations(s),"AC00000097")
				txtDF=GetMachineStationDefectCode(rsJ("JOB_NUMBER"),rsJ("SHEET_NUMBER"),astations(s),stationtotaldefectcodequantity)
				station_assembly_yield=csng(stationstartquantity-stationtotaldefectcodequantity)/csng(stationstartquantity)
				rsU("STATION_START_QUANTITY")=stationstartquantity
				rsU("STATION_DEFECTCODE_QUANTITY")=stationtotaldefectcodequantity
				rsU("STATION_ASSEMBLY_YIELD")=station_assembly_yield
				rsU.update
				else
				stationstartquantity=rsU("STATION_START_QUANTITY")
				txtDF=GetMachineStationDefectCode(rsJ("JOB_NUMBER"),rsJ("SHEET_NUMBER"),astations(s),stationtotaldefectcodequantity)
				stationtotaldefectcodequantity=rsU("STATION_DEFECTCODE_QUANTITY")
				station_assembly_yield=csng(rsU("STATION_ASSEMBLY_YIELD"))
				end if
			end if
			rsU.close
	%>
	<td><div align="center"><%=txtDF%></div></td>
    <td><div align="center"><% =formatdate(station_start_time,application("longdateformat"))%>&nbsp;</div></td>
	<td><div align="center"><%=formatpercent(station_assembly_yield,2,-1)%></div></td>
	<%		next%>
  </tr>
<%
		p=p+1
		rsJ.movenext
		wend
		end if
		rsJ.close
		end if
	i=i+1
	q=0
	p=1
	rs.movenext
	wend
else
%>
</table>
  <tr>
    <td height="20" colspan="10"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close
set rsJ=nothing
set rsU=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->