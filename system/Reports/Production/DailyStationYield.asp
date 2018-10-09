 
<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetSeries.asp" -->
<!--#include virtual="/Functions/GetFamily.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->

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

StationName=request("StationName")
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
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,j.part_number_tag as model_number"
		END IF
		
		if   model<>"" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,j.part_number_tag as model_number,js.Job_Number,js.Sheet_Number"
		end if
		
		SQL=SQL+", SUM(js.STATION_START_QUANTITY) AS INPUT, SUM(js.GOOD_QUANTITY) AS OUTPUT, "	
		SQL=SQL+" to_char(round((SUM(js.GOOD_QUANTITY) / decode(sum(js.STATION_START_QUANTITY),0,1,sum(js.STATION_START_QUANTITY))),4)*100) || '%' as Yield"
	
	
		if  family<>""  AND family<>"1" AND (isnull(Series) OR Series="")  then
			SQL=SQL+ ",to_char(MAX(FA.TARGET_FIRSTYIELD)) || '%'as target_FPY"
		END IF
		
		if  family<>""  AND family="1" then
			SQL=SQL+ ",to_char(MAX(SS.TARGET_FIRSTYIELD))|| '%' as target_FPY"
		END IF
		
		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
			SQL=SQL+ ",to_char(MAX(SN.TARGET_FIRSTYIELD))|| '%' as target_FPY"
		END IF
	
		if  SubSeries<>"" then
			SQL=SQL+ ",to_char(MAX(SS.TARGET_FIRSTYIELD))|| '%' as target_FPY"
		END IF
		SQL=SQL+",j.Job_Type"
		SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM, JOB j,JOB_STATIONS JS ,STATION S  "
   
		 
		SQL=SQL+" where j.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID  "
		SQL=SQL+" and j.job_number=js.job_number and j.sheet_number=js.sheet_number and JS.STATION_ID=S.NID"
		SQL=SQL+" AND JS.STATION_START_QUANTITY>0"
		
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
		
		SQL=SQL+" and js.close_time>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and js.close_time<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" AND S.MOTHER_STATION_ID='"+StationName+"'"
		
		SQL=SQL+" GROUP BY FAMILY_NAME "
		
		 
		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
			 
		end if
		
		if  family<>"" AND (isnull(Series) OR Series="")  and family<>"1" then
			SQL=SQL+ ",SERIES_NAME"
		 
		END IF
		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
			 
		END IF
		
		if  SubSeries<>"" AND (isnull(model) OR model="") then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,j.part_number_tag"
		END IF
		
		if  Model<>""  then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,j.part_number_tag,js.job_number, js.sheet_number"
		END IF
		SQL=SQL+",j.Job_Type"
		
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
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,j.part_number_tag"
		END IF
		
		if  Model<>""  then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME,j.part_number_tag,js.job_number,js.sheet_number"
		END IF
	 	SQL=SQL+",j.Job_Type"
	'pagepara="&seriesgroupname="&seriesgroupname&"&thisstatus="&thisstatus&"&factory="&factory
	
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
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script>
	function LoadNextItem()
	{
		form1.action="DailyStationYield.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="DailyStationYield.asp?Action=GenereateReport"
		form1.submit();
	}
	
	function GenerateDetailReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="DailyStationYield.asp?Action=GenereateDetialReport"
		form1.submit();
	}
</script>

</head>

<body onLoad="language(<%=session("language")%>);">
<form action="DailyStationYield.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="9" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr align="center">
   	<td height="20" width="80"><span id="inner_SearchFamily"></span></td>
     <td height="20" width="100">
	 	<select name="family" id="family" ONCHANGE="LoadNextItem()" style="width:100px">
    	<option value=""></option>
    	<%= getFamily("OPTION",family,factorywhereoutside,"","") %>
 		 </select>  
	</td> 
    <td width="80"><span id="inner_Series"></span></td>
     <td height="20" width="80"><select name="Series" id="Series" ONCHANGE="LoadNextItem()" style="width:100px">
    	<option value=""></option>
   		 <%= getSeries("OPTION",Series,wherestr1,"","") %>
 	 </select>  </td>
  
    <td width="100"><span id="inner_SubSeries"></span></td>
    <td height="20" width="80">
		<select name="SubSeries" id="SubSeries" ONCHANGE="LoadNextItem()" style="width:100px">
    		<option value=""></option>
	   		 <%= getSubSeries("OPTION",SubSeries,wherestr2,"","") %>
 	 	</select>  
	 </td>
   <Td width="80"><span id="inner_SearchPartNumber"></span> </Td>
  	 <td height="20" width="80" colspan="2" align="left">
	 	<select name="model" id="model" style="width:130px">
   			 <option value=""></option>
	   			 <%= getModel("OPTION",model,wherestr3,"","") %>
  		</select>  
	  </td>	   
  </tr>  
  <Tr align="center">
  	<td height="20" width="80"><span id="inner_StationName"></span></td>
  	<td height="20" width="200">
	 	<select name="StationName" id="StationName">
    		<%= getStation_New(true,"OPTION",StationName,factorywhereoutsideand," order by S.FACTORY_ID,S.STATION_NAME","","") %>
 		 </select>  
	</td>
  	 <td width="80"><span id="inner_CloseTime"></span></td> 
	 <td colspan="5" width="510" ><span id="inner_SearchFrom"></span>&nbsp; :
	 	<input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">	
        <script language=JavaScript type=text/javascript>
		function calendar1Callback(date, month, year)
		{
		document.all.fromdate.value=year + '-' + month + '-' + date
		}
		calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
							</script>
		<select name="fromtime" id="fromtime">
				 <option value="07:15:00" <% if fromtime="07:15:00" then response.write "Selected" end if%>>07:15:00</option>				
				 <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
				 <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>	 
  		</select>  
	&nbsp;<span id="inner_SearchTo"></span>:
	<input name="todate" type="text" id="todate" value="<%=todate%>" size="16">
	<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
	<select name="totime" id="totime">
				 <option value="07:15:00" <% if totime="07:15:00" then response.write "Selected" end if%>>07:15:00</option>	
   			 <option value="14:30:00" <% if totime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if totime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
		</td>
  	<Td align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="GenerateReport()"></Td>
  </Tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><span id="inner_BrowseData"></span></td>        
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right">
      <span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('../NewReport/Export_Excel.asp')"><img src="/Images/EXCEL.gif" width="22" height="22"></span></div>
	  </td>
    </tr>
  </table></td>
</tr>
<%if(rs.State > 0 ) then%>
  <tr>
  	<td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
	<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Family"></span></div></td>
	<%for i=0 to rs.Fields.count-1
	  spanId=""
	  if rs.Fields(i).name ="SERIES_NAME" then
		spanId="td_Series"
	  elseif rs.Fields(i).name ="SUBSERIES_NAME" then
		spanId="td_SubSeries"
	  elseif rs.Fields(i).name ="MODEL_NUMBER" then
		spanId="td_Model"
	  elseif rs.Fields(i).name ="JOB_NUMBER" then
		spanId="td_JobNumber"
	  elseif rs.Fields(i).name ="SHEET_NUMBER" then
		spanId="td_SheetNumber"
	  elseif rs.Fields(i).name ="TARGET_FPY" then
		targetYield="td_TargetYield"			  					
	  end if
	  if spanId <> "" then
	%> 
		<td height="20" class="t-t-Borrow"><div align="center"><span id='<%=spanId%>'></span></div></td>
  	<%end if
	next%>
	 <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Input"></span></div></td>	
	 <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Output"></span></div></td>
	 <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Yield"></span></div></td>
	 <%if targetYield <> "" then%>
	 <td height="20" class="t-t-Borrow"><div align="center"><span id='td_TargetYield'></span></div></td>
	 <%end if%>
	 <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_SearchJobType"></span></div></td>
	 <td height="20" class="t-t-Borrow"><div align="center"><span id="td_DefectInfo"></span></div></td>

	</tr>
	<%for j=0 to rs.recordcount-1
		input=input+cdbl(rs("INPUT"))
		moutput=moutput+cdbl(rs("OUTPUT"))
		mFamily=rs("FAMILY_NAME")		
	%>
	<tr align="center"><td><%=(j+1)%></td>
		<%for i=0 to rs.Fields.count-1%>
				 <td height="20" ><div align="center">				 
				 <%if(rs.Fields(i).name="SHEET_NUMBER" ) then
				 	jobNumber=rs("Job_Number")
					sheetNumber=rs("sheet_number")
					response.write "<a href='../../job/SubJobs/JobDetail.asp?jobnumber="&jobNumber&"&sheetnumber="&sheetNumber&"&jobtype="&rs("job_type")&"' target='_blank'>"
					end if
				 %>
				 <%=rs(i).value%>				 
				 <%if(rs.Fields(i).name="SHEET_NUMBER" ) then
						response.write "</a>"
					end if
				 %>				 
				  </div></td>
				 
				 <%
					if(rs.Fields(i).name="SERIES_NAME") then
							mSeries=rs(i).value
					elseif(rs.Fields(i).name="SUBSERIES_NAME") then
							mSubSeries=rs(i).value
					elseif(rs.Fields(i).name="MODEL_NUMBER") then
							mModel=rs(i).value
					end if				
				 %>
  		<%next%>
		 <td height="20"><div align="center">		 
		 <input type="button" id="btnDefectcode" name="btnDefectcode" value="Defect Info" onclick="window.open('DefectInfo.asp?Family=<%=mFamily%>&Series=<%=mSeries%>&SubSeries=<%=mSubSeries%>&Model=<%=mModel%>&Job_Number=<%=jobNumber%>&Sheet_Number=<%=sheetNumber%>&STATION=<%=StationName%>')">
		 </div></td>
	</tr>
	<%rs.movenext
	next 
	%>
  <tr>
   <tr align="center"><td class="t-t-Borrow"><span id="td_Total"></span></td>
   <%for i=0 to rs.Fields.count-1%>
	<% if(rs.Fields(i).name="INPUT") then %>
		<td height="20" class="t-t-Borrow"><div align="center"><%=input%></div></td>
	<% elseif(rs.Fields(i).name="OUTPUT") then %>
		<td height="20" class="t-t-Borrow"><div align="center"><%=moutput%></div></td>
	<% elseif(rs.Fields(i).name="YIELD") then %>
		<td height="20" class="t-t-Borrow"><div align="center">
				<% if input=0 then response.write "0" else response.write ROUND(moutput/input *100,2) %>%</div></td>
	<% else %>
	<td height="20" class="t-t-Borrow"><div align="center"> &nbsp;</div></td>
	<%end if
	next 
	%>
	 <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>	
   </tr>
<%end if%> 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->