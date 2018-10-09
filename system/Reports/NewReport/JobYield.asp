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
jobtype=request("jobtype")
JobNumber=""

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
		SQL="select  FA.FAMILY_NAME "
		
		SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag as model_number,Job_Number"
		
		
		SQL=SQL+", J.START_QUANTITY AS INPUT "
		
		'SQL=SQL+", (SELECT SUM(DEFECT_QUANTITY) FROM JOB_DEFECTCODES WHERE JOB_NUMBER=J.JOB_NUMBER) AS OUTPUT "	

		SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM, job_master j "
   
		 
		SQL=SQL+" where j.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID "
		
		
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
		
		SQL=SQL+" and j.INPUT_TIME>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and j.INPUT_TIME<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		
		
		'SQL=SQL+" GROUP BY FAMILY_NAME "
'		
'		 
'		if family="1" then
'			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
'		end if
'		
'		if  family<>"" AND (isnull(Series) OR Series="")  and family<>"1" then
'			SQL=SQL+ ",SERIES_NAME"
'		 
'		END IF
'		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
'			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
'		END IF
'		
'		if  SubSeries<>"" AND (isnull(model) OR model="") then
'			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag"
'		END IF
'		
'		if  Model<>""  then
'			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag,job_number"
'		END IF
		
		SQL=SQL +" ORDER BY FAMILY_NAME "
		SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,part_number_tag,job_number"
	pagename="/reports/newreport/yieldreport.asp"
	session("SQL")=SQL
	session("FromDateTime") =fromdate & " " &fromtime
	session("ToDateTime") =todate & " " &totime
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
	function LoadNextItem()
	{
		form1.action="JobYield.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="JobYield.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="JobYield.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Job Yield Report</td>
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
	 <td>&nbsp;</td>
	 <td>&nbsp;</td>
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
	%>
	 <td height="20" class="t-t-Borrow"><div align="center">Total Reject</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">Final Yield</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">Detail Info</div></td>
	<%
			
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
				 <%=rs(i).value%>
				  </div></td>
				 
				 <%
					if(rs.Fields(i).name="INPUT") then
						SingleInput=cdbl(rs(i).value)
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
				
				 %>
  	<%			
				next	
				
	%>
	 <%
	 		JobNumber=rs("Job_Number").value
			set rsReject=server.createobject("adodb.recordset")
			SQL="SELECT decode(SUM(DEFECT_QUANTITY),NULL,0,'',0,SUM(DEFECT_QUANTITY)) FROM JOB_DEFECTCODES JD,DEFECTCODE D "
			SQL=SQL+" WHERE JD.DEFECT_CODE_ID=D.NID AND D.TRANSACTION_TYPE='2' AND JD.JOB_NUMBER='"+JobNumber+"'"
	
			rsReject.open SQL,conn,1,3
			if rsReject.recordcount>0 then
				RejectCount=cdbl(rsReject(0).value)
				TotalRejectCount=TotalRejectCount+RejectCount
			else
				RejectCount=0
			end if
			
			if (SingleInput=0) then
				SingleInput=1
			end if
			FinalYield=round((SingleInput-RejectCount)/SingleInput*100,2)
			
		%>
				 
	<td height="20" ><div align="center"><%=cstr(RejectCount)%></div></td>
    <td height="20" ><div align="center"><%=cstr(FinalYield)%>%</div></td>
	 <td height="20"><div align="center">
	  
	 <input type="button" id="btnDefectcode" name="btnDefectcode" value="Detail Info" onclick="window.open('JobDetailYield.asp?Job_Number=<%=rs("Job_Number").value%>')">
	 
	 </div></td>
	<%
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
				<%end if%>
  	<%     
		    next 
	%>
	 <td height="20" class="t-t-Borrow"><div align="center"><%=TotalRejectCount%></div></td>
	 <%if input=0 then input=1 end if%>
	 <td height="20" class="t-t-Borrow"><div align="center"><%=round((input-TotalRejectCount)/input*100,2)%>%</div></td>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
	<%
		end if
	%>
   </tr>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->