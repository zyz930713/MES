<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->

<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
StationName=request("StationName")
reportstationid=""

partnumber=request("partnumber")
linename=request("linename")

fromdate=request("fromdate")
fromtime=request("fromtime")
time0=now 
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(dateadd("d",-weekday(time0)+1,time0))) +"-"+cstr(month(dateadd("d",-weekday(time0)+1,time0)))+"-"+cstr(day(dateadd("d",-weekday(time0)+1,time0)))
end if

if isnull(fromtime) or fromtime=""  then
	fromtime="14:30:00"
end if

todate=request("todate")
if isnull(todate) or todate=""  then
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
totime=request("totime")
if isnull(totime) or totime="" then
	totime="14:30:00"
end if
reportstationid=StationName
if(StationName="SA00000222") then
	sqlstr=GetStackMagnetAssemblySQL(reportstationid,fromdate+" "+fromtime,todate+" "+totime,partnumber,linename)
end if 

if(StationName="SA00000161") then
	sqlstr=GetCSYMASQL(reportstationid,fromdate+" "+fromtime,todate+" "+totime,partnumber,linename)
end if 

if(StationName="SA00000163") then
	sqlstr=GetReedWeldSQL(reportstationid,fromdate+" "+fromtime,todate+" "+totime,partnumber,linename)
end if 

if(StationName<>"" and sqlstr<>"") then
	session("SQL")=sqlstr
	rs.open sqlstr,conn,1,3
end if 
%>


<%
	function GetStackMagnetAssemblySQL(reportstationid,starttime,endtime,partnumber,linename)
		 sql="select a.job_number,a.sheet_number,a.operator_code,a.close_time,a.station_start_quantity, a.good_quantity,"
		 sql=sql+ "d.part_number_tag,d.line_name,"
		 sql=sql+ "("
		 sql=sql+ " select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
		 sql=sql+ " where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
		 sql=sql+ "  and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"' and cc.nid='AN00000557'"
		 sql=sql+ " ) as MagnetScrapQty, "
		 sql=sql+ " ("
		 sql=sql+ "  select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
		 sql=sql+ " where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
         sql=sql+ " and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"' and cc.nid='AN00000263'"
		 sql=sql+" )as StackScrapQty"
		 sql=sql+" from job_stations a,job d,station b,station_new c"
         sql=sql+" where a.station_id=b.nid and b.mother_station_id=c.nid and c.nid='"+reportstationid+"'"
		 sql=sql+" and a.job_number= d.job_number and a.sheet_number=d.sheet_number" 
		  if(partnumber<>"") then
			 sql=sql+" and d.part_number_tag='"+partnumber+"'"
		 end if 
		 if(linename<>"") then
			 sql=sql+" and d.line_name='"+linename+"'"
		 end if 
		 if(starttime<>"") then
			 sql=sql+" and a.close_time>to_date('"+starttime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 if(endtime<>"") then
			 sql=sql+" and a.close_time<to_date('"+endtime+"','yyyy-mm-dd hh24:mi:ss')"
		 end if 
		 GetStackMagnetAssemblySQL=sql
	end function 
	
	function GetCSYMASQL(reportstationid,starttime,endtime,partnumber,linename)
			sql="select a.job_number,a.sheet_number,a.operator_code,"
			sql=sql+" (select operator_code from job_stations aa, station ss where aa.job_number=a.job_number and aa.sheet_number=a.sheet_number and aa.station_id=ss.nid and ss.mother_station_id='SA00000222') as CSMA_OP,"
			sql=sql+" a.close_time,a.station_start_quantity, a.good_quantity,"
			sql=sql+ "d.part_number_tag,d.line_name,"
 			sql=sql+" ("
			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
			sql=sql+" and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"'"
			sql=sql+" and cc.nid='AN00000497'"
			sql=sql+" ) as SteponMagnetStack, "
			sql=sql+" ("
			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
			sql=sql+" and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"'"
			sql=sql+" and cc.nid='AN00000517'"
			sql=sql+" )as CoilNGScrapQty "
			sql=sql+" from job_stations a,job d,station b,station_new c"
			sql=sql+" where a.station_id=b.nid and b.mother_station_id=c.nid and c.nid='"+reportstationid+"'"
			sql=sql+" and a.job_number= d.job_number and a.sheet_number=d.sheet_number" 
			
			 if(partnumber<>"") then
				 sql=sql+" and d.part_number_tag='"+partnumber+"'"
			 end if 
			 if(linename<>"") then
				 sql=sql+" and d.line_name='"+linename+"'"
			 end if 
		 
			 if(starttime<>"") then
				 sql=sql+" and a.close_time>to_date('"+starttime+"','yyyy-mm-dd hh24:mi:ss')"
			 end if 
			 if(endtime<>"") then
				 sql=sql+" and a.close_time<to_date('"+endtime+"','yyyy-mm-dd hh24:mi:ss')"
			 end if 
			 GetCSYMASQL=sql
	end function 
	
	function GetReedWeldSQL(reportstationid,starttime,endtime ,partnumber,linename)
			
			sql="select a.job_number,a.sheet_number,a.operator_code,"
			sql=sql+" (select operator_code from job_stations aa, station ss where aa.job_number=a.job_number and aa.sheet_number=a.sheet_number and aa.station_id=ss.nid and ss.mother_station_id='SA00000161') as Stack_OP,"
			sql=sql+" a.close_time,a.station_start_quantity, a.good_quantity,"
			sql=sql+ "d.part_number_tag,d.line_name,"
			sql=sql+"("
			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
			sql=sql+"  and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"'"
			sql=sql+" and cc.nid='AN00000244'"
			sql=sql+" ) as MachineCode, "
			sql=sql+" ("
			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
			sql=sql+"  and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"' "
			sql=sql+" and cc.nid='AN00000558'"
			sql=sql+" )as ReedScrapQty, "
			sql=sql+" ("
			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
			sql=sql+"  and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"'"
			sql=sql+" and cc.nid='AN00000537'"
			sql=sql+" ) as Bad_CSMA_C,"
			sql=sql+" ("
			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
  			sql=sql+" and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"'"
			sql=sql+" and cc.nid='AN00000518'"
			sql=sql+") as Bad_CSMA_F,"
			sql=sql+"("
 			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
  			sql=sql+" and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"'"
			sql=sql+" and cc.nid='AN00000538'"
			sql=sql+") as Bad_CSMA_Angle,"
			sql=sql+"("
 			sql=sql+" select aa.action_value from job_actions aa, action bb, action_new cc,station dd,station_new ee"
			sql=sql+" where aa.action_id=bb.nid and bb.mother_action_id=cc.nid and aa.station_id=dd.nid and dd.mother_station_id=ee.nid"
  			sql=sql+" and aa.job_number= a.job_number and aa.sheet_number=a.sheet_number and ee.nid='"+reportstationid+"'"
			sql=sql+" and cc.nid='AN00000225'"
			sql=sql+" ) as Coil_fall_off"
			sql=sql+" from job_stations a,job d,station b,station_new c"
			sql=sql+" where a.station_id=b.nid and b.mother_station_id=c.nid and c.nid='"+reportstationid+"'"
			sql=sql+" and a.job_number= d.job_number and a.sheet_number=d.sheet_number" 
			 if(partnumber<>"") then
				 sql=sql+" and d.part_number_tag='"+partnumber+"'"
			 end if 
			 if(linename<>"") then
				 sql=sql+" and d.line_name='"+linename+"'"
			 end if 
		 
			 if(starttime<>"") then
				 sql=sql+" and a.close_time>=to_date('"+starttime+"','yyyy-mm-dd hh24:mi:ss')"
			 end if 
			 if(endtime<>"") then
				 sql=sql+" and a.close_time<=to_date('"+endtime+"','yyyy-mm-dd hh24:mi:ss')"
			 end if 
			 GetReedWeldSQL=sql
	 
	end function
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
	function GenerateReport()
	{
		form1.action="GEYieldReprt.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="GEYieldReprt.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">GE Yield Report</td>
  </tr>
  <tr>
   <td height="20">Station Name(Õ¾Ãû)</td>
  	<td height="20">
	 	<select name="StationName" id="StationName">
    		<%= getStation_New(true,"OPTION",StationName,factorywhereoutsideand," order by S.FACTORY_ID,S.STATION_NAME","","") %>
 		 </select>  
	</td>
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
	   		 <option value="05:00:00" <% if fromtime="05:00:00" then response.write "Selected" end if%>>05:00:00</option> 
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
			 <option value="05:00:00" <% if fromtime="05:00:00" then response.write "Selected" end if%>>05:00:00</option>
	   			 
  		</select>  
	</td>
  	<Td>&nbsp;</Td>
	<Td>&nbsp;</Td>
  </Tr>
  <tr> 
  	<Td>PartNumber</Td>
	<Td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>" size="16"></Td>
	<Td>LineName</Td>
	<Td><input name="LineName" type="text" id="LineName" value="<%=partnumber%>" size="16"></Td>
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
	<%
		if(rs.State > 0 ) then
			for i=0 to rs.Fields.count-1
	%>
    <td height="20" class="t-t-Borrow"><div align="center"><%=rs.Fields(i).name%> </div></td>
  	<%
			next 
		end if
	%>
</tr>
	<%  if(rs.State > 0 ) then	
			for j=0 to rs.recordcount-1
	%>
	<tr>
	<%			
		for i=0 to rs.Fields.count-1
	%>
   			 <td height="20" ><div align="center"><% if isnull(rs(i).value) then response.write "0" else response.write rs(i).value end if %> </div></td>
  	<%			
				next
				rs.movenext
			next 
	%>
		</tr>
	<%
		end if
	%>
	
  </tr>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->