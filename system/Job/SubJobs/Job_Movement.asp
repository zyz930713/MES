<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
Response.Buffer = False

%>
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetJobStation.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<%
	currentstation=trim(request("currentstation"))
	fromdate=request("txtJobCloseTimeFrom")
	todate=request("txtJobCloseTimeTo")
	jobclosetimefromhour=request("txtJobCloseTimeFromHour")
	if(jobclosetimefromhour="") then
		jobclosetimefromhour="00:00:00"
		
	end if
	JobCloseTimeToHour=request("txtJobCloseTimeToHour")
	if(JobCloseTimeToHour="") then
		JobCloseTimeToHour="23:59:59"
	end if
	linename=request("txtLineName")
	
	timeflag=request("dropTimeFlag")
 	RESPONSE.WRITE timeflat
	
	SQL="select PART_NUMBER_TAG,LINE_NAME,station_name,sum(JS.Station_Start_quantity) as  Station_Start_quantity ,sum(JS.good_quantity ) as good_quantity "
	SQL=SQL & "from job_stations JS,JOb j,station S"
	SQL=SQL &" where JS.job_number=j.job_number and js.sheet_number=j.sheet_number and JS.station_id=s.NID"
	IF(currentstation <> "") then
		SQL=SQL & " AND S.MOTHER_STATION_ID='"&currentstation&"'"
	end if
	IF(fromdate <> "") then
		if(timeflag="" or timeflag="0") then
			SQL=SQL & " AND JS.close_time>=TO_DATE('"& fromdate &" " & jobclosetimefromhour & "','YYYY-MM-DD HH24:MI:SS')"
		else
			SQL=SQL & " AND JS.start_time>=TO_DATE('"& fromdate &" " & jobclosetimefromhour & "','YYYY-MM-DD HH24:MI:SS')"
		end if
		session("starttime")=fromdate &" " & jobclosetimefromhour
		session("timeflag")=timeflag
	end if
	
	IF(todate <> "") then
		if(timeflag="" or timeflag="0") then
			SQL=SQL & " AND JS.close_time<=TO_DATE('"& todate &" "& JobCloseTimeToHour & "','YYYY-MM-DD HH24:MI:SS')"
		else
			SQL=SQL & " AND JS.start_time<=TO_DATE('"& todate &" "& JobCloseTimeToHour & "','YYYY-MM-DD HH24:MI:SS')"
		end if
		session("endtime")=todate &" " & JobCloseTimeToHour
		session("timeflag")=timeflag
	end if
	if(linename<>"") then
		SQL=SQL & " AND J.line_name='"&linename&"'"
	end if
	 if(request.querystring("action")="") then
		SQL=SQL & " AND 1=2 group by PART_NUMBER_TAG,LINE_NAME,station_name "
		rs.open SQL,conn,1,3
		
	 end if
	
 	 if(request.querystring("action")="Query") then
		SQL=SQL & " group by PART_NUMBER_TAG,LINE_NAME,station_name order by  PART_NUMBER_TAG,LINE_NAME,station_name"
		 rs.open SQL,conn,1,3
	 end if
	 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<title><%=application("SystemName")%></title>
	<link href="/CSS/General.css" rel="stylesheet" type="text/css">
	<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
	<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
	<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
	 <script>
	 	function tableexpand(i,url,count)
		{
			if(eval("document.all.tabimg"+i+".title=='Expand'"))
			{
			eval("document.all['tab"+i+"'].style.display=''");
			eval("document.all.tabimg"+i+".src='/Images/Treeimg/minus.gif'");
			eval("document.all.tabimg"+i+".title='Collapse'");
			eval("document.all.tabfrm"+i+".src='"+url+"'")
				if (count<10)
				{eval("document.all.tabfrm"+i+".height="+(count+2)*20)}
			}
			else
			{
			eval("document.all['tab"+i+"'].style.display='none'");
			eval("document.all.tabimg"+i+".src='/Images/Treeimg/plus.gif'");
			eval("document.all.tabimg"+i+".title='Expand'");
			eval("document.all.tabfrm"+i+".src=''")
			}
		}


	 </script>
</head>
<body>
<form action="" method="post" name="form1">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_StationName">Job Movement Report</span></td>
    </tr>
    <tr>
      <td>Station Name</td>
      <td><select name="currentstation" id="currentstation">
        <option value="">All</option>
		 <%=getStation_New(true,"OPTION",currentstation,""," order by S.STATION_NAME","","")%>
      </select></td>
	  
      <td>Station <select name="dropTimeFlag" id="dropTimeFlag">
          <option value="0" <% if timeflag="0" THEN response.write "SELECTED" END IF%>>Close</option>
          <option value="1"  <% if timeflag="1" THEN response.write "SELECTED" END IF%>>Start</option>
       </select> 
	    Time From</td>
      <td  height="20"><input name="txtJobCloseTimeFrom"  type="text" value="<%=fromdate%>" >
	   <script language=JavaScript type=text/javascript>
		function calendar1Callback(date, month, year)
		{
			document.all.txtJobCloseTimeFrom.value=year + '-' + month + '-' + date
		}
		calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
	  </script>
	  <input name="txtJobCloseTimeFromHour"  type="text" value="<%=jobclosetimefromhour%>" >
	  </td>
	  <td>To</td>
	  <td  height="20"><input name="txtJobCloseTimeTo"  type="text" value="<%=todate%>" >
	  <script language=JavaScript type=text/javascript>
		function calendar2Callback(date, month, year)
		{
			document.all.txtJobCloseTimeTo.value=year + '-' + month + '-' + date
		}
		calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
	  </script>
	  <input name="txtJobCloseTimeToHour"  type="text" value="<%=JobCloseTimeToHour%>" >
	  <input type="button" name="btnQuery" onclick="form1.action='job_movement.asp?action=Query';form1.submit();" value="Query" width="75px" class="t-b-Yellow" >
	  <input type="button" name="btnExport" onclick="form1.action='Job_Movement_Excel.asp';form1.submit();" value="Export" width="75px" class="t-b-Yellow" >
	  <input type="button" name="btnExport_Detail" onclick="form1.action='Job_Movement_Detail_Excel.asp';form1.submit();" value="Export Detail" width="75px" class="t-b-Yellow" >
	  </td>
    </tr>
	<Tr>
		 <td>Line Name</td>
	   <td><input name="txtLineName"  type="text" value="<%=LineName%>" ></td>
		<td></td>
		<td></td>
		<td></td>
		<td></td>
	</Tr>
	<Tr>
	<td colspan="6" width="100%">
		<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
		<tr>
 
			<Td height="20" class="t-t-Borrow">Part Number</Td>
			<Td height="20" class="t-t-Borrow">Line Name</Td>
			<Td height="20" class="t-t-Borrow">Station Name</Td>
 
			<Td height="20" class="t-t-Borrow">Station Start Quantity</Td>
			<Td height="20" class="t-t-Borrow">Good Quantity</Td>
		</tr>
		<% dim i
			i=1
			while not rs.eof %>
			<tr>
				<td>
				<span style="cursor:hand" onClick="tableexpand(<%=i%>,'/Job/SubJobs/Job_MovementDetail.asp?PartNumber=<%=rs(0).value%>&LineName=<%=rs(1).value%>&currentstation=<%=currentstation%>',20)">
				<img src="/Images/Treeimg/plus.gif" name="tabimg<%=i%>" width="9" height="9" border="0" align="absmiddle" id="tabimg<%=i%>" title="Expand">
				</span>
				<%=rs(0).value%>
				</td>
				<td><%=rs(1).value%></td>
				<td><%=rs(2).value%></td>
				<td><%=rs(3).value%></td>
				<td><%=rs(4).value%></td>
			
			</tr>
			<tbody id="tab<%=i%>" style="display:none">
			  <tr>
				<td colspan="5"><iframe src="" width="100%" height="200" scrolling="auto" frameborder="0" id="tabfrm<%=i%>"></iframe></td>
			  </tr>
			</tbody>
		 <% 
		  i=i+1
		 rs.movenext 
		 wend %>
	  </table>
	 </td>
	 </Tr>
	</table>
	
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
