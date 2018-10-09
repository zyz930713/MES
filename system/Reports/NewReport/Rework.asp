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
adjustcount=0
reworkcount=0
SCRAPcount=0

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


if action="GenereateReport" then
		SQL="select  FA.FAMILY_NAME "

		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		end if
		if  family<>"" and family<>"1" AND (isnull(Series) OR Series="")  then
			SQL=SQL+ ",SN.SERIES_NAME"
		END IF
		
		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		END IF
		
		if  SubSeries<>""  then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag as model_number"
		END IF
		
		SQL=SQL+", decode(transaction_type,'1' ,'Rework','2','Scrap','3', 'Adjust' ,'0','none') as transaction_type ,SUM(defect_quantity) AS defect_quantity "	
		
		SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM, job j,job_defectcodes jd,  defectcode df "
   
		 
		SQL=SQL+" where j.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID  "
		SQL=SQL+" and j.job_number = jd.job_number AND j.sheet_number = jd.sheet_number  AND jd.defect_code_id = df.nid"
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
		
		SQL=SQL+" GROUP BY transaction_type,FAMILY_NAME"',SERIES_NAME,SUBSERIES_NAME,part_number_tag"
		
		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		end if
		
		if  family<>"" AND (isnull(Series) OR Series="")  and family<>"1" then
			SQL=SQL+ ",SERIES_NAME"
		END IF
		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		END IF
		
		if  SubSeries<>""  then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag"
		END IF
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
	 
	 FinalSQL="select distinct family_name,series_name,subseries_name,model_number,( select defect_quantity from("+SQL+")TEMP1 where temp1.model_number=temp.model_number and temp1.transaction_type='Rework') as rework,"
	 FinalSQL=FinalSQL+" ( select defect_quantity from("+SQL+")TEMP1 where temp1.model_number=temp.model_number and temp1.transaction_type='Scrap') as Scrap, "
	 FinalSQL=FinalSQL+" ( select defect_quantity from("+SQL+")TEMP1 where temp1.model_number=temp.model_number and temp1.transaction_type='Adjust') as Adjust "
	 FinalSQL=FinalSQL+" FROM ("+SQL+") TEMP"
	pagename="/reports/newreport/FinalReport.asp"
	'pagepara="&seriesgroupname="&seriesgroupname&"&thisstatus="&thisstatus&"&factory="&factory
	
	session("SQL")=FinalSQL
	'response.write SQL
	'response.end

	rs.open FinalSQL,conn,1,3
	
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
		form1.action="Rework.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}
		if(form1.family.options(form1.Series.selectedIndex).value=="0")
		{
			window.alert("Please select one Series!");
			return;
		}
		if(form1.family.options(form1.SubSeries.selectedIndex).value=="0")
		{
			window.alert("Please select one SubSeries!");
			return;
		}

		form1.action="Rework.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="Rework.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Rework Summary Report</td>
  </tr>
  <tr>
    <td height="20">Family  Name </td>
     <td height="20">
	 	<select name="family" id="family" ONCHANGE="LoadNextItem()">
    	<option value="0">-- Select Family --</option> 
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
	 <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
	
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
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
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
	<Td></Td>
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
   			 <td height="20" ><div align="center"><% if isnull(rs(i).value)  then response.write "0" else response.write rs(i).value end if%></div></td>
			 <%
			 	if(rs.Fields(i).name="REWORK") then
					if(not isnull(rs(i).value))  then
						reworkcount=reworkcount+cdbl(rs(i).value)
					end if
				end if
				
				if(rs.Fields(i).name="SCRAP") then
					if(not isnull(rs(i).value))  then
						SCRAPcount=SCRAPcount+cdbl(rs(i).value)
					end if
				end if
				
				if(rs.Fields(i).name="ADJUST") then
					if(not isnull(rs(i).value))  then
						adjustcount=adjustcount+cdbl(rs(i).value)
					end if
				end if
			 %>
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
   <tr>
   <%
		if(rs.State > 0 ) then
			for i=0 to rs.Fields.count-1
	%>
				<% if(rs.Fields(i).name="REWORK") then %>
					<td height="20" class="t-t-Borrow"><div align="center"><%=reworkcount%></div></td>
				<% elseif(rs.Fields(i).name="SCRAP") then %>
					<td height="20" class="t-t-Borrow"><div align="center"><%=SCRAPcount%></div></td>
				<% elseif(rs.Fields(i).name="ADJUST") then %>
					<td height="20" class="t-t-Borrow"><div align="center">
							<%=adjustcount%></td>
				<% else %>
				<td height="20" class="t-t-Borrow"><div align="center"> &nbsp;</div></td>
				<%end if%>
  	<%     
		    next 
		end if
	%>
   </tr>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->