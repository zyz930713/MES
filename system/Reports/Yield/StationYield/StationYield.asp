<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetJobActionValue.asp" -->
<!--#include virtual="/Functions/GetMachineStationDefectCode.asp" -->
<!--#include virtual="/Functions/GetJobTotalDefectCodeQuantity.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rsU=server.CreateObject("adodb.recordset")
set rsV=server.CreateObject("adodb.recordset")
station=request("station")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if fromdate="" then
fromdate=dateadd("d",-1,date())
end if
if todate="" then
todate=date()
end if

if station<>"" and station<>"all" then
uwhere=" and L.NID='"&station&"'"
end if
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER='"&jobnumber&"'"
end if
if partnumber<>"" then
where=where&" and P.PART_NUMBER='"&partnumber&"'"
end if
if fromdate<>"" then
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
else
fromdate=date()
where=where&" and J.START_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.START_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&station="&station&"&jobnumber="&jobnumber&"&partnumber="&partnumber&"&fromdate="&fromdate&"&todate="&todate&"&ordername="&ordername&"&ordertype="&ordertype
pagename="/Reports/Yield/StationYield/StationYield.asp"
SQL="select 1,J.JOB_NUMBER,J.SHEET_NUMBER,J.START_TIME,J.PART_NUMBER_TAG,P.PART_NUMBER from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER is not null"&where&order
rs.open SQL,conn,1,3
FactoryRight "L."
SQLU="select 1,L.NID,L.STATION_NAME from STATION L inner join FACTORY F on L.FACTORY_ID=F.NID where L.STATUS=1"&uwhere&factorywhereoutsideand
rsU.open SQLU,conn,1,3
Tcount=rsU.recordcount
session("SQL")=SQL
session("where")=where
session("order")=order

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
<div id="RefreshDIV" style="position:absolute;visibility:hidden;"><iframe id="RefreshFrame" src=""></iframe></div>
<form name="form1" method="post" action="/Reports/Yield/StationYield/StationYield.asp">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-c-greenCopy"><span>Search Station Yield</span></td>
  </tr>
  <tr>
    <td>Station</td>
    <td><select name="station" id="station">
      <option value="">All</option>
	  <%FactoryRight "S."%>
      <%=getStation(true,"OPTION",station,factorywhereoutside," order by STATION_NAME","","")%>
    </select></td>
    <td height="20"><span class="style1">Job Number</span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td>Part Number </td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td>Job Start Time</td>
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
  <td height="20" colspan="<%=Tcount*2+4%>" class="t-c-greenCopy">Browse Station Yield </td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount*2+4%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%">User:
        <% =session("User") %></td>
      <td width="50%"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('StationYield_Export.asp')"><img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount*2+4%>"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&station=<%=station%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&fromdate=<%=fromdate%>&todate=<%=todate%>'">Job Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&station=<%=station%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&station=<%=station%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&fromdate=<%=fromdate%>&todate=<%=todate%>'">Part Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=P.PART_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&station=<%=station%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&station=<%=station%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&fromdate=<%=fromdate%>&todate=<%=todate%>'">Start Time<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.START_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%>&station=<%=station%>&jobnumber=<%=jobnumber%>&partnumber=<%=partnumber%>&fromdate=<%=fromdate%>&todate=<%=todate%>'"> </div></td>
  <%while not rsU.eof%>
  <td class="t-t-Borrow"><div align="center"><%=rsU("STATION_NAME")%>&nbsp;</div></td>
  <%rsU.movenext
  wend
  rsU.movefirst%>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize
%>
<tr>
  	<td height="20" rowspan="<%=q%>"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td height="20"><div align="center"><a href="/Job/SubJobs/JobDetail.asp?jobnumber=<%=rs("JOB_NUMBER")%>" target="_blank"><%= rs("JOB_NUMBER") %>-<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%></a></div></td>
    <td height="20"><div align="center"><%= rs("PART_NUMBER_TAG") %>&nbsp;</div></td>
    <td><div align="center"><% =formatdate(rs("START_TIME"),application("longdateformat"))%></div></td>
	<%
	while not rsU.eof
		stationstartquantity=0
		stationtotaldefectcodequantity=0
		station_assembly_yield=0
		txtDF="&nbsp;"
		SQLV="select STATION_START_QUANTITY,STATION_DEFECTCODE_QUANTITY,STATION_ASSEMBLY_YIELD from JOB_STATIONS where JOB_NUMBER='"&rs("JOB_NUMBER")&"' and SHEET_NUMBER='"&rs("SHEET_NUMBER")&"' and STATION_ID='"&rsU("NID")&"'"
		rsV.open SQLV,conn,1,3
		if not rsV.eof then
			if cint(rsV("STATION_ASSEMBLY_YIELD"))=0 then
			stationstartquantity=getJobActionValue(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rsU("NID"),"AC00000097")
			txtDF=GetMachineStationDefectCode(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rsU("NID"),stationtotaldefectcodequantity)
			station_assembly_yield=csng(stationstartquantity-stationtotaldefectcodequantity)/csng(stationstartquantity)
			rsV("STATION_START_QUANTITY")=stationstartquantity
			rsV("STATION_DEFECTCODE_QUANTITY")=stationtotaldefectcodequantity
			rsV("STATION_ASSEMBLY_YIELD")=station_assembly_yield
			rsV.update
			else
			stationstartquantity=rsV("STATION_START_QUANTITY")
			txtDF=GetMachineStationDefectCode(rs("JOB_NUMBER"),rs("SHEET_NUMBER"),rsU("NID"),stationtotaldefectcodequantity)
			stationtotaldefectcodequantity=rsV("STATION_DEFECTCODE_QUANTITY")
			station_assembly_yield=csng(rsV("STATION_ASSEMBLY_YIELD"))
			end if
		end if
		rsV.close
	%>
    <td><span id="y<%=rs("JOB_NUMBER")%>_<%=repeatstring(rs("SHEET_NUMBER"),"0",3)%>_<%=rsU("NID")%>"><table border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><%=txtDF%></td>
  </tr>
  <tr>
    <td><div align="center"><%=formatpercent(station_assembly_yield,2,-1)%><%if station_assembly_yield<>0 then%><img src="/Images/Refresh1.gif" alt="Click button to refresh value" width="15" height="15" align="absmiddle" style="cursor:hand" onClick="javascript:document.all.RefreshFrame.src='StationYield_Refresh.asp?jobnumber=<%=rs("JOB_NUMBER")%>&sheetnumber=<%=rs("SHEET_NUMBER")%>&station=<%=rsU("NID")%>'"><%end if%></div></td>
  </tr>
</table></span></td>
    <%rsU.movenext
	wend
	rsU.movefirst%>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="<%=Tcount*2+4%>"><div align="center">No Records </div></td>
  </tr>
<%end if
rsU.close
rs.close
set rsU=nothing
set rsv=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->