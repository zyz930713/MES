<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
StationName=request("StationName")
action=request.querystring("Action")
jobnumber=request("jobnumber")
sheetunmber=request("sheetunmber")
dayout=request("dayout")
if(dayout=null or dayout="") then
	'dayout=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
'------------------------------------------Report------------------------------------------------------------------
if action="GenereateReport" then
		
		if (stationname="SA00000141") then
			SQL=SQL+" SELECT 工单号,小工单号,型号,站名,中文站名,工单起始数量,工单当前数量 ,工单开始时间,工单类型,线别,拍打放置天数,出站时间 FROM ("
		end if
			SQL=SQL+" select j.job_number as 工单号, j.sheet_number as 小工单号,j.part_number_tag as 型号,sn.station_name as 站名,sn.station_chinese_name as 中文站名, "
			SQL=SQL+" j.job_start_quantity as 工单起始数量,j.job_good_quantity as 工单当前数量,j.start_time as 工单开始时间,j.job_type as 工单类型,j.line_name as 线别"
			
			if (stationname="SA00000141") then
				SQL=SQL+", (select decode(c.SlappingHoldDays,null,0,'',0,c.SlappingHoldDays)  as Slapping_Hold_Time"
				SQL=SQL+"  FROM General_Setting  c where c.ModelName=j.part_number_tag) as 拍打放置天数"
				
				SQL=SQL+",(start_time+(select decode(c.SlappingHoldDays,null,0,'',0,c.SlappingHoldDays)  as Slapping_Hold_Time FROM  General_Setting  c "
				SQL=SQL+" where c.ModelName=j.part_number_tag)) as 出站时间 "
			end if
			SQL=SQL+" from job j ,station s,station_new sn"
			SQL=SQL+" where j.current_station_id=s.nid and s.mother_station_id=sn.nid"
			SQL=SQL+" and j.status='0' and j.close_time is null "
			SQL=SQL+" and sn.is_delete=0 "
			
			if(jobnumber<>"")then
				SQL=SQL+" and j.job_number='"+jobnumber+"'"
			end if
			
			if(sheetunmber<>"")then
				SQL=SQL+" and j.sheet_number='"+sheetunmber+"'"
			end if
			
			
			if(StationName<>"1")then
				SQL=SQL+" and sn.nid='"+StationName+"'"
			end if
		if (stationname="SA00000141") then
			SQL=SQL+") TEMP "
			if(dayout<>"")then
				SQL=SQL+" WHERE 出站时间>=to_date(to_char(to_date('"+dayout+"','yyyy-mm-dd'),'yyyy-mm-dd')||'00:00:00','yyyy-mm-dd hh24:mi:ss')"
				SQL=SQL+" AND 出站时间<=to_date(to_char(to_date('"+dayout+"','yyyy-mm-dd'),'yyyy-mm-dd')||'23:59:59','yyyy-mm-dd hh24:mi:ss')"
			end if
			SQL=SQL+" ORDER BY 出站时间 DESC"
		ELSE
			SQL=SQL+" order by j.current_station_id,j.part_number_tag,j.job_number,sheet_number"
		end if	
		session("SQL")=SQL
		'response.write sql
		'RESPONSE.END
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
	 
	function GenerateReport()
	{
		 
		form1.action="WIPReport.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="WIPReport.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">WIP Report</td>
  </tr>
  <tr>
    <td height="20">Station Name(站名)</td>
     <td height="20">
	 	<select name="StationName" id="StationName">
		<option value="1" <%if StationName="1" then response.write "Selected" end if%>>All Station</option>
    		<%= getStation_New(true,"OPTION",StationName,factorywhereoutsideand," order by S.FACTORY_ID,S.STATION_NAME","","") %>
 		 </select>  
	</td>
  		<td height="20">
			Job Number(工单号)
	    </td>
	   <Td>
	   		<input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">
	   </Td>
  	 	<td height="20">
	 	    Sheet Number(小工单号)
	    </td>
	  	<td height="20">
			<input name="sheetnumber" type="text" id="sheetnumber" value="<%=sheetnumber%>">
	    </td>
  	<Td>出打寄存站时间<input name="dayout" type="text" id="dayout" value="<%=dayout%>">
		<script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.dayout.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
	</Td>
	 <Td>	<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()">
<span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="20" height="20"></span></Td>
    
  </tr>
   
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
				 <td height="20" ><div align="center">
				 
				 <%
				 	if(rs.Fields(i).name="SHEET_NUMBER" or rs.Fields(i).name="JOB_NUMBER") then
						response.write "<a href='../../job/SubJobs/JobDetail.asp?jobnumber="&rs("Job_Number").value&"&sheetnumber="&rs("sheet_number").value&"&jobtype="&rs("job_type").value&"' target='_blank'>"
					end if
				 %>
				 <%=rs(i).value%> &nbsp;
				 
				 <%
				 	if(rs.Fields(i).name="SHEET_NUMBER" or rs.Fields(i).name="JOB_NUMBER") then
						response.write "</a>"
					end if
				 %>
				 
				  </div></td>
				 
				 <%
					if(rs.Fields(i).name="INPUT") then
						input=input+cdbl(rs(i).value)
					end if
					
					if(rs.Fields(i).name="OUTPUT") then
						moutput=moutput+cdbl(rs(i).value)
					end if
					
				if(rs.Fields(i).name="FAMILY_NAME") then
						mFamily=rs(i).value
				end if
				if(rs.Fields(i).name="SERIES_NAME") then
						mSeries=rs(i).value
				end if
				if(rs.Fields(i).name="SUBSERIES_NAME") then
						mSubSeries=rs(i).value
				end if
				if(rs.Fields(i).name="MODEL_NUMBER") then
						mModel=rs(i).value
				end if
				next	
				rs.movenext
			next 
	%>
		</tr>
	<%
		end if
	%>
	
  </tr>
   <tr>
   <%
		if(rs.State > 0 ) then
			for i=0 to rs.Fields.count-1
	%>
				<% if(rs.Fields(i).name="INPUT") then %>
					<td height="20" class="t-t-Borrow"><div align="center"><%=input%></div></td>
				<% elseif(rs.Fields(i).name="OUTPUT") then %>
					<td height="20" class="t-t-Borrow"><div align="center"><%=moutput%></div></td>
				<% elseif(rs.Fields(i).name="FPY") then %>
					<td height="20" class="t-t-Borrow"><div align="center">
							<% if input=0 then response.write "0" else response.write ROUND(moutput/input *100,2) %>%</div></td>
				<% else %>
				<td height="20" class="t-t-Borrow"><div align="center"> &nbsp;</div></td>
				<%end if
		    next 
		end if
	%>
   </tr>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->