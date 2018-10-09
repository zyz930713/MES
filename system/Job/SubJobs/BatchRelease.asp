<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
 
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Functions/GetSeriesGroup.asp" -->
<!--#include virtual="/Functions/GetStationTransactionChange.asp" -->
<!--#include virtual="/Functions/GetStationOperator.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/Functions/NID_SEQ.asp" -->
<!--#include virtual="/Functions/HoldReleaseFunc.asp" -->
<%
jobnumber=trim(request("jobnumber"))
sheetnumber=trim(request("sheetnumber"))
holdType=trim(request("hold_type"))
partnumber=trim(request("partnumber"))
timespan=trim(request("timespan"))
jobstatus=trim(request("jobstatus"))
line=request("line")
currentstation=trim(request("currentstation"))
fromdate=request("fromdate")
if fromdate = "" then
	fromdate=dateadd("d",-7,date())
end if
todate=request("todate")
close_fromdate=request("close_fromdate")
close_todate=request("close_todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
	order=" order by H.TRANSACTION_TIME "
else
	order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
	where=where&" and H.JOB_NUMBER = '"&jobnumber&"'"
end if
if sheetnumber<>"" then
	where=where&" and H.SHEET_NUMBER='"&sheetnumber&"'"
end if
if holdType<>"" then
	where=where&" and H.HOLD_TYPE_ID='"&holdType&"'"
end if
if partnumber<>"" then
	where=where&" and J.PART_NUMBER_TAG = '"&partnumber&"'"
end if
if line<>"" then
	where=where&" and lower(J.LINE_NAME) like '%"&lcase(line)&"%'"
end if

if currentstation<>"" then
	where=where&" and S.MOTHER_STATION_ID='"&currentstation&"'"
end if

if fromdate<>"" then
	where=where&" and H.TRANSACTION_TIME>=to_date('"&fromdate&"','yyyy-mm-dd') "
end if
if todate<>"" then
	where=where&" and H.TRANSACTION_TIME<=to_date('"&todate&"','yyyy-mm-dd')"
end if
if (request.querystring("Action")="") then
	where=where&" and 1=2"
end if 

if(request.querystring("Action")="Hold") then 
	jobcount=request("jobcount")
	reason=request("HoldReason")
	jobNum = ""
	partNum = "" 
	holdIdList = ""
	for m=0 to jobcount-1 
		if (request("checkbox"&m)="on") then
			job_number=request("sjobnumber"&m)
			sheet_number=request("ssheetnumber"&m)
			station_id=request("sstationid"&m)
			holdId=request("hold_id"&m)
		 	set rsHoldJob1=server.createobject("adodb.recordset")
			SQL="select CONTROL_TYPE,CONTROL_STATION,CONTROL_REASON,CONTROL_PERSON,CONTROL_TIME,STATUS,PART_NUMBER_TAG from JOB where JOB_NUMBER='"&job_number&"' and SHEET_NUMBER='"&sheet_number&"'"
			rsHoldJob1.open SQL,conn,1,3
			rsHoldJob1("CONTROL_TYPE")=""	
			rsHoldJob1("CONTROL_STATION")=""
			rsHoldJob1("CONTROL_REASON")=""
			rsHoldJob1("CONTROL_PERSON")=""		
			rsHoldJob1("CONTROL_TIME")=""				
			rsHoldJob1("STATUS")=0
			partNum = partNum & rsHoldJob1("PART_NUMBER_TAG") & ";"
			rsHoldJob1.update
			rsHoldJob1.close
				
			holdIdList = holdIdList & "," & holdId
			releaseId=cstr(NID_SEQ("TRANSACTION"))
			
			SQL="UPDATE JOB_HOLD_RELEASE_HISTORY SET RELATED_ID='"+releaseId+"' WHERE ID='"+holdId+"' "
			rs.open SQL,conn,3,3
			
			SQL="insert into JOB_HOLD_RELEASE_HISTORY (JOB_NUMBER,SHEET_NUMBER,STATION_ID,TRANSACTION_TYPE,"
			SQL=SQL+" TRANSACTION_PERSON,TRANSACTION_TIME,TRANSACTION_REASON,ID,RELATED_ID)"
			SQL=SQL+"VALUES('"+job_number+"','"+sheet_number+"','"+station_id+"','0',"
			SQL=SQL+" '"+session("code")+"','"+CSTR(now())+"','"+reason+"','"
			sql=sql+releaseId+"','"
			sql=sql+holdId+"')"
			
			set rsHoldJob2=server.createobject("adodb.recordset")
			rsHoldJob2.open SQL,conn,1,3
			
			'add by Lennie Wei 2013-01-10
            jobNum = jobNum  & job_number & "-" & repeatstring(sheet_number,"0",3) & ";"
		end if 
	next 
	'add by Lennie Wei 2013-01-10     
     sendHoldReleaseNotiInfo jobNum,"Release",reason,holdIdList,partNum
    'end add by Lennie Wei 2013-01-10
    
	RESPONSE.WRITE("<SCRIPT>window.alert('Release job successfully!')</SCRIPT>")
end if 

SQL="SELECT H.JOB_NUMBER,H.SHEET_NUMBER,H.TRANSACTION_TIME,H.TRANSACTION_REASON,H.ID,B.GROUP_NAME|| '-' || B.GROUP_CHINESE_NAME AS HOLD_TYPE,B.GROUP_MEMBERS,J.PART_NUMBER_TAG,J.JOB_START_QUANTITY, "
SQL=SQL+" (SELECT PART_NUMBER FROM PART WHERE NID=J.PART_NUMBER_ID) AS PART_NUMBER,STATION_NAME||'-('||STATION_CHINESE_NAME||')' AS STATION,J.START_TIME,J.CURRENT_STATION_ID,J.LINE_NAME "
SQL=SQL+" FROM JOB_HOLD_RELEASE_HISTORY H INNER JOIN SYSTEM_GROUP B ON H.HOLD_TYPE_ID=B.NID INNER JOIN JOB J ON H.JOB_NUMBER=J.JOB_NUMBER AND H.SHEET_NUMBER=J.SHEET_NUMBER "
SQL=SQL+" INNER JOIN STATION S ON H.STATION_ID = S.NID WHERE H.RELATED_ID IS NULL AND J.STATUS=2 "
SQL=SQL+ where&order

rs.open SQL,conn,1,3
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
<script language=javascript src="/Language/language.js" type=text/javascript></script>
<script>

    function Query()
   {
		    form1.action="BatchRelease.asp?Action=Query";
			form1.submit();
   }

	function HoldJob()
	{
		var jobcount=document.getElementById("jobcount").value;
		var i=0;
		var selectjobcount=0
		for(i=0;i<jobcount;i++)
		{
			if(document.getElementById("checkbox"+i).checked==true)
			{
				selectjobcount++;
			}
		}
		if(selectjobcount==0)
		{
			window.alert("请选择要Release的Job!");
			return;
		}
		else
		{
			if (document.getElementById("HoldReason").value=="")
			{
				window.alert("请输入Reason!");
				return;
			}
			form1.action="BatchRelease.asp?Action=Hold";			
			form1.btnOK.disabled=true;
			form1.submit();
		}
	}
</script>
</head>

<body onLoad="language('<%=session("language")%>');form1.btnOK.disabled=false;">
<form  method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_ReleaseJob"></span></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_SearchJobNumber"></span></td>
      <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>">
        -
          <input name="sheetnumber" type="text" id="sheetnumber" value="<%=sheetnumber%>" size="3" maxlength="3"></td>
      <td height="20"><span id="inner_SearchLineName"></span></td>
      <td height="20"><input name="line" type="text" id="line" value="<%=line%>"></td>
      <td><span id="inner_SearchCurrentStation"></span></td>
      <td><select name="currentstation" id="currentstation">
        <option value="">All</option>
        <%=getStation_New(true,"OPTION",currentstation,""," order by S.STATION_NAME","","")%>
      </select></td>      
    </tr>  
    <tr>
      <td><span id="inner_SearchHoldType"></span></td>
      <td>
	  	<select name="hold_type" id="hold_type">
		<option value="">All</option>
		<%=getHoldType(holdType)%>
		</select>   
      </td>
      <td><span id="inner_SearchPartNumber"></span></td>
      <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
       <td height="20"><span id="inner_SearchHoldTime"></span></td>
      <td height="20"><span id="inner_SearchFrom"></span><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
        <script language=JavaScript type=text/javascript>
				function calendar1Callback(date, month, year)
				{
				document.all.fromdate.value=year + '-' + month + '-' + date
				}
				calendar1 = new dynCalendar('calendar1', 'calendar1Callback', document.all.fromdate.value);
				calendar1.setMonthCombo(true);
				calendar1.setYearCombo(true);
			</script>
			&nbsp;<span id="inner_SearchTo"></span>
			<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
			<script language=JavaScript type=text/javascript>
				function calendar2Callback(date, month, year)
				{
				document.all.todate.value=year + '-' + month + '-' + date
				}
				calendar2 = new dynCalendar('calendar2', 'calendar2Callback', document.all.todate.value);
          </script>
&nbsp;
<img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onclick="Query()">
</td>
    </tr>
    
</table>

  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  	<tr>
		<td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
	</tr>   
    <tr>
      <td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_User"></span>:
      <% =session("User") %></td>
    </tr>
     
    <tr>
	  <td height="20" class="t-t-Borrow"><div align="center">&nbsp;</div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_JobNumber"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_PartNumber"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_PartType"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Line"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_Quantity"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_StartTime"></span></div></td>
      <td class="t-t-Borrow"><div align="center"><span id="inner_HoldTime"></span></div></td>
	  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_CurrentStations"></span></div></td>
	  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_HoldType"></span></div></td>
      <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_HoldReason"></span></div></td>
    </tr>
<%
i=0
while not rs.eof 
	sjobnumber=rs("JOB_NUMBER")&"-"&repeatstring(rs("SHEET_NUMBER"),"0",3)
	groupMembers=rs("GROUP_MEMBERS")	
	current_station_id=rs("CURRENT_STATION_ID")
%>
    <tr>
	  <td height="20"><div align="center">
	   	<input type="checkbox" name="checkbox<%=i%>" id="checkbox<%=i%>"
		<%if groupMembers<>"" and instr(lcase(groupMembers),lcase(session("email")))=0 then response.Write("disabled") end if%>	
		>
	    </div></td>
      <td height="20"><div align="center"><%=sjobnumber%><input type="hidden" name="sjobnumber<%=i%>" id="sjobnumber<%=i%>" value='<%=rs("JOB_NUMBER")%>'>
	  													<input type="hidden" name="ssheetnumber<%=i%>" id="ssheetnumber<%=i%>" value='<%=rs("SHEET_NUMBER")%>'</div></td>
      <td><div align="center"><%=rs("PART_NUMBER_TAG")%></div></td>
      <td><div align="center"><%=rs("PART_NUMBER")%></div></td>
      <td><div align="center"><%=rs("LINE_NAME")%></div></td>
      <td><div align="center"><%=rs("JOB_START_QUANTITY")%></td>
      <td height="20"><div align="center">
        <% =formatdate(rs("START_TIME"),application("longdateformat"))%>
        &nbsp;</div></td>
      <td><div align="center"><%=formatdate(rs("TRANSACTION_TIME"),application("longdateformat"))%>&nbsp;</div></td>
	  <td><%=rs("STATION")%></td>
      <td><div align="center"><%=rs("HOLD_TYPE")%>&nbsp;</div></td>
      <td><div align="center"><%=rs("TRANSACTION_REASON")%>&nbsp;</div></td>
	  <input type="hidden" name="sstationid<%=i%>" id="sstationid<%=i%>" value="<%=current_station_id%>" />
	  <input type="hidden" name="hold_id<%=i%>" id="hold_id<%=i%>" value="<%=rs("ID")%>" />
    </tr>
<%
i=i+1
rs.movenext
wend
rs.close%>

<tr>
      <td height="20" colspan="16">
	  	<Table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<Tr>
				<td colspan="2"  class="t-c-greenCopy">&nbsp;</td>
			</Tr>
			<Tr>
				<td width="10%" valign="top"><span id="inner_ReleaseReason"></span></td>
				<td width="90%"><textarea name="HoldReason" id="HoldReason" cols="100" rows="10"><%=reason%></textarea></td>
			</Tr>
			<Tr>
				<td colspan="2"><input type="button" id="btnOK" name="btnOK" onclick="HoldJob()" value="OK"></td>
				 
			</Tr>
		</Table>
	  </td>
</tr>
</table>
<input type="hidden" id="jobcount" name="jobcount" value=<%=i%>>
</form>
  <!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
