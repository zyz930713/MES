<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->


<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
family=request("family")
Series=request("Series")
SubSeries=request("SubSeries")
model=request("model")
action=request("Action")
fromdate=request("fromdate")
fromtime=request("fromtime")
input=0
moutput=0
time0=now   
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
end if

if isnull(fromtime) or fromtime=""  then
	fromtime="14:30:00"
end if
todate=request("todate")
if isnull(todate) or todate=""  then
	'todate=cstr(year(dateadd("d",7-weekday(time0),time0)))+"-"+cstr(month(dateadd("d",7-weekday(time0),time0)))+"-"+cstr(day(dateadd("d",7-weekday(time0),time0)))
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
totime=request("totime")
if isnull(totime) or totime="" then
	totime="14:30:00"
end if
wherestr1=" where 1=2"
wherestr2=" where 1=2"
wherestr3=" where 1=2"

if not isnull(family) and family<>"" then
	 if(action="LoadNextItem" OR action="GenereateReport" ) then
	 	wherestr1=" where Family_id='"&family&"'"
	 end if
end if

if not isnull(Series) and Series<>"" then
	if(action="LoadNextItem" OR action="GenereateReport" ) then
	 	wherestr2=" where Series_id='"&Series&"'"
	 end if
end if

if  not isnull(SubSeries) and SubSeries<>"" then
	if(action="LoadNextItem" OR action="GenereateReport") then
	 	wherestr3=" where Series_id='"&SubSeries&"'"
	 end if
end if

'------------------------------------------Summary Report------------------------------------------------------------------
if action="GenereateReport" then
		SQL="select /*+ ordered use_hash(b,a) */  FA.FAMILY_NAME "

		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		end if
		if  family<>"" and family<>"1" AND (isnull(Series) OR Series="")  then
			SQL=SQL+ ",SN.SERIES_NAME"
		END IF
		
		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		END IF
		
		if  SubSeries<>"" AND (isnull(Model) OR Model="")  then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag as model_number"
		END IF
		
		if  model<>"" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag as model_number"
		end if
		
		SQL=SQL+", a.job_number,a.sheet_number, j.line_name,j.part_number_tag ,b.station_group_name, a.station_id,a.uvs  "

		SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM, job j ,job_station_group_uvs a ,  station_group b"
   
		 
		SQL=SQL+" where j.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID  "
		SQL=SQL+" and  a.job_number= j.job_number and a.sheet_number=j.sheet_number and a.station_group_id= b.station_group_id and b.status='1'  "
		
		
		if not isnull(family) and family<>"" and family<>"1" then
			SQL=SQL+" and PM.family_id='"&family&"'"
			
			 
		END IF
		
		if  NOT isnull(Series) AND Series<>"" then
			SQL=SQL+" and PM.SERIES_GROUP_ID='"&Series&"' "
			 
		END IF
		
		if  NOT isnull(SubSeries) AND SubSeries<>"" then
			SQL=SQL+" and PM.SERIES_ID='"&SubSeries&"' "
			 
		END IF
		
		if  NOT isnull(model) AND model<>"" then
			SQL=SQL+" and PM.ITEM_NAME='"& model &"' "
		 
		END IF
		
		SQL=SQL+" and j.close_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and j.close_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and j.job_type='N'"
		'SQL=SQL+" and a.job_number='7957684' and a.sheet_number='1'"
		

		SQL=SQL +" ORDER BY FAMILY_NAME "
		
		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		end if
		
		
		if  family<>"" AND (isnull(Series) OR Series="") and family<>"1" then
			SQL=SQL+ ",SERIES_NAME"
		END IF
		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		END IF
		
		if  SubSeries<>"" AND (isnull(model) OR model="") then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag"
		END IF
		
		if  Model<>""  then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag,job_number,  sheet_number"
		END IF
	 	SQL=SQL+",a.sequence_no"
	rs.open SQL,conn,1,3
	
end if
 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script>
	function LoadNextItem()
	{
		form1.action="StationGroupFinanceYield.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="StationGroupFinanceYield.asp?Action=GenereateReport"
		form1.submit();
	}
	
	function GenerateDetailReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="StationGroupFinanceYield.asp?Action=GenereateDetialReport"
		form1.submit();
	}

	function ShowJobInfo()
	{
		window.open("../../job/subjobs/Job_Export_Excel.asp");
	}
</script>

</head>

<body>
<form action="StationGroupFinanceYield.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Station Group Finance Yield</td>
  </tr>
  <tr>
    <td height="20">Family  Name </td>
     <td height="20">
	 	<select name="family" id="family" ONCHANGE="LoadNextItem()">
    	<option value="0">-- Select Family --</option>
		<option value="1" <%if family="1" then response.write "Selected" end if%>>All Family</option>
 
    	<%= getFamily("OPTION",family,factorywhereoutside,"","") %>
 		 </select>  
	</td>
  
  
    <td>Series Name</td>
     <td height="20"><select name="Series" id="Series" ONCHANGE="LoadNextItem()">
    	<option value="">-- Select Series --</option>
   		 <%= getSeries("OPTION",Series,wherestr1,"","") %>
 	 </select>  </td>
  
    <td>Sub Series Name</td>
    <td height="20">
		<select name="SubSeries" id="SubSeries" ONCHANGE="LoadNextItem()">
    		<option value="">-- Select Sub Series --</option>
	   		 <%= getSubSeries("OPTION",SubSeries,wherestr2,"","") %>
 	 	</select>  
	 </td>
   <Td>Model </Td>
  	 <td height="20">
	 	<select name="model" id="model">
   			 <option value="">-- Select Model --</option>
	   			 <%= getModel("OPTION",model,wherestr3,"","") %>
  		</select>  
	  </td>
  </tr>
  
  <Tr>
  	 <td>DateTime From:</td> 
	 <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">
	
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
	<select name="fromtime" id="fromtime">
   			 <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
</td>
<td>To:</td>
<td>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="16">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
	<select name="totime" id="totime">
   			 <option value="14:30:00" <% if totime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if totime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
	
&nbsp;
  	<Td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>
  	<Td><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span></Td>
	<Td>&nbsp;</Td>
	 <Td>&nbsp;</Td>
  </Tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="18"> </td>
  </tr>
  <tr>
    <td height="20" class="t-t-Borrow"><div align="center">Job Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Sheet Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Part Number</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Line Name</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Station Group</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">OutPut Qty</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">UVS</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">OP NO</div></td>
	<td height="20" class="t-t-Borrow"><div align="center">Output Hour</div></td>
  	
</tr>
	<%if action="GenereateReport" then %>
	
	<%			
				while not rs.eof
	%>
		<tr>
				 <td height="20" ><%=rs("job_number")%></td>
				 <td height="20" ><%=rs("sheet_number")%> </td>
				 <td height="20" ><%=rs("part_number_tag")%> </td>
				 <td height="20" ><%=rs("line_name")%></td>
				 <td height="20" ><%=rs("station_group_name")%> </td>
				  <%
				  	stations_index=rs("STATION_ID")
					stationsArr=split(stations_index,",")
					stationstr=""
					for i=0 to ubound(stationsArr)
						if(stationsArr(i)<>"")then
							stationstr=stationstr+"'"+stationsArr(i)+"',"
						end if 
					next 
					if(len(stationstr)>0) then
						stationstr=left(stationstr,len(stationstr)-1)
					end if 
				 	subsqlstr="select OPERATOR_CODE,GOOD_QUANTITY from JOB_STATIONS where job_number='"+rs("job_number")+"' and sheet_number='"+rs("sheet_number")+"'"
					subsqlstr=subsqlstr+" and station_id in("+stationstr+")   order by close_time"
		 
					set rsStationInfo=server.createobject("adodb.recordset")
					rsStationInfo.open subsqlstr,conn,1,3
					opcode=""
					qty="0"
					while not rsStationInfo.eof
						opcode=opcode+rsStationInfo("OPERATOR_CODE")+","
						qty=rsStationInfo("GOOD_QUANTITY")
						rsStationInfo.movenext
					wend
				 %>
				 
				 <td height="20" ><%=qty%></td>
				 <td height="20" ><%=rs("UVS")%> </td>
				 <td height="20" ><%=opcode%> </td>
				 <%
				 	
					 OutputHour=cdbl(qty)*cdbl(rs("UVS"))
				 %>
				 <td height="20" ><%=OutputHour%></td>
			</tr>
  	<%			
					rs.movenext
				wend
	%>
	
	<%end if %> 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->