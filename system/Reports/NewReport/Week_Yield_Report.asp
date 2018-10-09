<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/POCF_Open.asp" -->
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
myear=request("syear")
mStartWeek=request("sStartWeek")
mEndWeek=request("sEndWeek")

input=0
moutput=0


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
		fromdatetime =get_Start_DateTime_New(mStartWeek,myear) 
		todatetime=get_End_DateTime_New(mEndWeek,myear)  
		SQL="select  FA.FAMILY_NAME "
		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		end if
		
		if  NOT isnull(Series) AND Series<>"" then
			SQL=SQL+ ",SN.SERIES_NAME"
		END IF
		
		if  NOT isnull(SubSeries) AND SubSeries<>"" then
			SQL=SQL+ ",SUBSERIES_NAME"
		END IF
		
		if  NOT isnull(model) AND model<>"" then
			SQL=SQL+ ",part_number_tag as model_number"
		END IF
		
		
		SQL=SQL+", SUM(jms.INPUT_QUANTITY) as Input,SUM(jms.STORE_QUANTITY) AS OutPut, "	
		SQL=SQL+" to_char(round((SUM(jms.STORE_QUANTITY) /  decode(SUM(jms.input_quantity),0,1,SUM(jms.input_quantity))),   4) *100) || '%' as FPY"
		
		if  family<>""AND family="1"  then
			SQL=SQL+ ",MAX(SS.TARGET_FIRSTYIELD) as target_FPY,"
		END IF
		
		if  family<>"" AND family<>"1" AND (isnull(Series) OR Series="")  then
			SQL=SQL+ ",MAX(FA.TARGET_FIRSTYIELD) as target_FPY,"
		END IF
		
		if  Series<>"" AND (isnull(SubSeries) OR SubSeries="") then
			SQL=SQL+ ",MAX(SN.TARGET_FIRSTYIELD) as target_FPY,"
		END IF
		
		if  SubSeries<>"" then
			SQL=SQL+ ",MAX(SS.TARGET_FIRSTYIELD) as target_FPY,"
		END IF
		
		SQL=SQL+" get_week_sequence(JMS.STORE_TIME) AS WW"
		 SQL=SQL+" from SERIES_NEW SN,SUBSERIES SS,FAMILY FA,PRODUCT_MODEL PM,JOB_MASTER_STORE JMS "
		SQL=SQL+" where JMS.part_number_tag=pm.item_name(+) and SN.NID=PM.SERIES_GROUP_ID AND SS.NID=PM.SERIES_ID AND FA.NID=PM.FAMILY_ID and JMS.STORE_TYPE='N' "
		
		SQL=SQL+" and JMS.JOB_NUMBER not like 'E%' and JMS.JOB_NUMBER not like '%-R' and JMS.JOB_NUMBER not like '%-Q' "
		SQL=SQL+" and JMS.JOB_NUMBER not like '%-E' and JMS.JOB_NUMBER not like '%-T' and JMS.JOB_NUMBER not like '%-D'  "
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
		
		SQL=SQL+" and JMS.STORE_TIME>to_date('"& fromdatetime &"','yyyy-mm-dd hh24:mi:ss')"
		SQL=SQL+" and JMS.STORE_TIME<=to_date('"& todatetime &"','yyyy-mm-dd hh24:mi:ss')"
		
		
		SQL=SQL+" GROUP BY FAMILY_NAME"
		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		end if
		
		if  NOT isnull(Series) AND Series<>"" then
			SQL=SQL+ ",SERIES_NAME"
		END IF
		if  NOT isnull(SubSeries) AND SubSeries<>"" then
			SQL=SQL+ ",SUBSERIES_NAME"
		END IF
		
		if  NOT isnull(model) AND model<>"" then
			SQL=SQL+ ",part_number_tag"
		END IF
		
		SQL=SQL+",get_week_sequence(JMS.STORE_TIME)"
		SQL=SQL +" ORDER BY FAMILY_NAME "
		if family="1" then
			SQL=SQL+ ",SERIES_NAME,SUBSERIES_NAME"
		end if
		if  NOT isnull(Series) AND Series<>"" then
			SQL=SQL+ ",SERIES_NAME"
		END IF
		if  NOT isnull(SubSeries) AND SubSeries<>"" then
			SQL=SQL+ ",SUBSERIES_NAME"
		END IF
		
		if  NOT isnull(model) AND model<>"" then
			SQL=SQL+ ",part_number_tag"
		END IF
		
		SQL=SQL+",get_week_sequence(JMS.STORE_TIME)"
	 
	pagename="/reports/newreport/yieldreport.asp"
	'pagepara="&seriesgroupname="&seriesgroupname&"&thisstatus="&thisstatus&"&factory="&factory
	
	session("SQL")=SQL
	'RESPONSE.WRITE sql
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
<script>
	function LoadNextItem()
	{
		form1.action="Week_Yield_Report.asp?Action=LoadNextItem"
		form1.submit();
	}
	
	function GenerateReport()
	{
		if(form1.family.options(form1.family.selectedIndex).value=="0")
		{
			window.alert("Please select one family!");
			return;
		}

		form1.action="Week_Yield_Report.asp?Action=GenereateReport"
		form1.submit();
	}
</script>

</head>

<body>
<form action="Week_Yield_Report.asp" method="post" name="form1" target="_self" id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Weekly Yield Report</td>
  </tr>
  <tr>
    <td height="20">Family  Name </td>
     <td height="20">
	 	<select name="family" id="family" ONCHANGE="LoadNextItem()"  >
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
  	 <td>Year:</td>
	 <td>
	 <select name="syear" id="syear" style="width:110px">
   			 <option value="2010" <%if myear="2010" then response.write "SELECTED" END IF%>>2010</option>
	   		 <option value="2011" <%if myear="2011" then response.write "SELECTED" END IF%>>2011</option> 
			 <option value="2012" <%if myear="2012" then response.write "SELECTED" END IF%>>2012</option> 
			 <option value="2013" <%if myear="2013" then response.write "SELECTED" END IF%>>2013</option> 
			 <option value="2014" <%if myear="2014" then response.write "SELECTED" END IF%>>2014</option> 
			 <option value="2015" <%if myear="2015" then response.write "SELECTED" END IF%>>2015</option> 
			 <option value="2016" <%if myear="2016" then response.write "SELECTED" END IF%>>2016</option> 
			 <option value="2017" <%if myear="2017" then response.write "SELECTED" END IF%>>2017</option> 
			 <option value="2018" <%if myear="2018" then response.write "SELECTED" END IF%>>2018</option> 
  	 </select> 
	 </td>
	 <td> 
	 	From:
		<select name="sStartWeek" id="sStartWeek">
			<%for mm=1 to 53 %>
   			 <option value="<%=mm%>" <%if cstr(mStartWeek)=cstr(mm) then response.write "SELECTED" END IF%>>WW<%if mm<10 then response.write "0"&mm else response.write mm end if%></option>
			<%next %>
      	</select> 
	</td>
	 <td> 
		&nbsp;To:
		<select name="sEndWeek" id="sEndWeek">
			<%for nn=1 to 53 %>
   			 <option value="<%=nn%>" <%if cstr(mEndWeek)=cstr(nn) then response.write "SELECTED" END IF%>>WW<%if nn<10 then response.write "0"&nn else response.write nn end if%></option>
			<%next %>
      	</select> 
		
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
   			 <td height="20" ><div align="center"><%=rs(i).value%> </div></td>
			 
			 <%
			 	if(rs.Fields(i).name="INPUT") then
					input=input+cdbl(rs(i).value)
				end if
				
				if(rs.Fields(i).name="OUTPUT") then
					moutput=moutput+cdbl(rs(i).value)
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
				<% if(rs.Fields(i).name="FAMILY_NAME") then %>
					<td height="20" class="t-t-Borrow"><div align="center"><%response.write "Total:"%></div></td>
				<% elseif(rs.Fields(i).name="INPUT") then %>
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
		end if
	%>
   </tr>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/POCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->