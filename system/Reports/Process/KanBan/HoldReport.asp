<%@  language="VBSCRIPT" codepage="936" %>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
action=request("Action")
fromdate=request("fromdate")
fromtime=request("fromtime")
jobnumber=request("jobnumber")
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


'------------------------------------------Summary Report------------------------------------------------------------------
if action="GenereateReport" then
	sql="SELECT JOB_NUMBER,SHEET_NUMBER,MODELNAME,LINE_NAME,HOLD_TIME,RELEASE_TIME,"
	sql=sql+"round((to_date(RELEASE_TIME,'yyyy-mm-dd hh24:mi:ss') - to_date(HOLD_TIME,'yyyy-mm-dd hh24:mi:ss'))*24,2)||'h' as PSS_TIME,"
	sql=sql+"(SELECT USER_NAME ||'('||USER_CODE||')' FROM USERS WHERE USER_CODE = A.HOLD_PERSON) AS HOLD_PERSON,"
	sql=sql+"(SELECT USER_NAME ||'('||USER_CODE||')' FROM USERS WHERE USER_CODE = A.RELEASE_PERSON) AS RELEASE_PERSON,"
	sql=sql+" (SELECT GROUP_NAME || '-' || GROUP_CHINESE_NAME FROM SYSTEM_GROUP WHERE NID=A.HOLD_TYPE_ID) AS HOLD_TYPE, "
	sql=sql+"HOLD_REASON,OP_CODE,"
	sql=sql+"(SELECT STATION_NAME ||'-'|| STATION_CHINESE_NAME FROM STATION WHERE NID=A.STATION_ID) AS HOLD_STATION "
	sql=sql+"FROM( "
	sql=sql+"select JHRH.JOB_NUMBER,JHRH.SHEET_NUMBER,J.PART_NUMBER_TAG AS MODELNAME,J.LINE_NAME,JHRH.TRANSACTION_TIME AS HOLD_TIME,JHRH.transaction_person AS HOLD_PERSON,"
	sql=sql+"(SELECT TRANSACTION_TIME FROM  JOB_HOLD_RELEASE_HISTORY A WHERE A.RELATED_ID=JHRH.ID) AS RELEASE_TIME,"
	sql=sql+" (SELECT TRANSACTION_PERSON FROM  JOB_HOLD_RELEASE_HISTORY A WHERE A.RELATED_ID=JHRH.ID) AS RELEASE_PERSON,"
	sql=sql+" (SELECT OPERATOR_CODE FROM JOB_STATIONS JS WHERE JS.JOB_NUMBER=J.JOB_NUMBER AND JS.SHEET_NUMBER=J.SHEET_NUMBER "
	sql=sql+" AND JS.STATION_ID=JHRH.STATION_ID) AS OP_CODE,"
	sql=sql+" STATION_ID,JHRH.TRANSACTION_REASON AS HOLD_REASON,HOLD_TYPE_ID, "
	sql=sql+" (SELECT GROUP_NAME || '-' || GROUP_CHINESE_NAME FROM SYSTEM_GROUP WHERE NID=JHRH.HOLD_TYPE_ID) AS HOLD_TYPE "
	sql=sql+" from JOB J, JOB_HOLD_RELEASE_HISTORY JHRH "
	sql=sql+" WHERE J.JOB_NUMBER=JHRH.JOB_NUMBER AND J.SHEET_NUMBER=JHRH.SHEET_NUMBER "
	sql=sql+" AND JHRH.TRANSACTION_TYPE='1'"
	
	SQL=SQL+" and JHRH.TRANSACTION_TIME>to_date('"& fromdate & " " &fromtime &"','yyyy-mm-dd hh24:mi:ss')"
	SQL=SQL+" and JHRH.TRANSACTION_TIME<=to_date('"& todate & " " &totime &"','yyyy-mm-dd hh24:mi:ss')"
		
	if jobnumber<>"" then
		SQL=SQL+" and j.job_number='"+jobnumber+"'"
	end if 
	sql=sql+") A ORDER BY JOB_NUMBER,HOLD_TIME DESC "
	
	session("SQL")=SQL
	'response.write sql	
	
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
	<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
    <script language="javascript" src="/Components/sniffer.js" type="text/javascript"></script>
    <script language="javascript" src="/Components/dynCalendar.js" type="text/javascript"></script>
    <script type="text/javascript">
        function GenerateReport() {
            form1.action = "HoldReport.asp?Action=GenereateReport"
            form1.submit();
        }
    </script>
</head >
<body onLoad="language(<%=session("language")%>);">
    <form action="HoldReport.asp" method="post" name="form1" target="_self" id="form1">
    <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
        <tr>
            <td height="20" colspan="8" class="t-c-greenCopy"><span id="inner_HoldReport"></span>
            </td>
        </tr>
        <tr align="center">
            <td width="80">
                <span id="inner_SearchJobNumber"></span>
            </td>
            <td width="100">
                <input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>" size="16">
            </td>
            <td width="100"><span id="inner_HoldTime"></span>&nbsp;<span id="inner_SearchFrom"></span>
            </td>
            <td width="200">
                <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">
                <script language="JavaScript" type="text/javascript">
                    function calendar1Callback(date, month, year) {
                        document.all.fromdate.value = year + '-' + month + '-' + date
                    }
                    calendar1 = new dynCalendar('calendar1', 'calendar1Callback', document.all.fromdate.value);
                </script>

                <select name="fromtime" id="fromtime">
                    <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>
                        14:30:00</option>
                    <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>
                        23:59:59</option>
                </select>
            </td>
            <td width="30">
                <span id="inner_SearchTo"></span>
            </td>
            <td width="200">
                <input name="todate" type="text" id="todate" value="<%=todate%>" size="16">
                <script language="JavaScript" type="text/javascript">
                    function calendar2Callback(date, month, year) {
                        document.all.todate.value = year + '-' + month + '-' + date
                    }
                    calendar2 = new dynCalendar('calendar2', 'calendar2Callback', document.all.todate.value);
                </script>
                <select name="totime" id="totime">
                    <option value="14:30:00" <% if totime="14:30:30" then response.write "Selected" end if%>>
                        14:30:00</option>
                    <option value="23:59:59" <% if totime="23:59:59" then response.write "Selected" end if%>>
                        23:59:59</option>
                </select>
                
                <td align="left">
                    <img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle"
                        style="cursor: hand" onclick="GenerateReport()" >
                &nbsp;
                    
                </td>
        </tr>
		<tr>
			<td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
		  </tr>
		  <tr>
			<td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>:
				  <% =session("User") %>        </td>	
				<td align="right">
					<span style="cursor: hand" title="Export Data to EXCEL File" onclick="<%if(rs.State > 0 ) then %>javascript:window.open('Export_Excel.asp')<%else %>return false;<%end if %>" > 
                        <img src="/Images/EXCEL.gif" width="20" height="20" >
					</span>
					&nbsp;
				</td>  			
			  </tr>
			</table></td>
		  </tr>		
    </table>
    </form>
    <%if(rs.State > 0 ) then %>
    <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">        
        <tr>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_JobNumber"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_SheetNumber"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_Model"></span></div></td>	
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_LineName"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_HoldTime"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_ReleaseTime"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_PSSTime"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_HoldPerson"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_ReleasePerson"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_HoldType"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_HoldReason"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_OpCode"></span></div></td>			
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_HoldStation"></span></div></td>			           
        </tr>
        <%for j=0 to rs.recordcount-1%>
        <tr>
            <%for i=0 to rs.Fields.count-1%>
            <td height="20"><div align="center">
                <%
                if rs(i).value <> "" then
                    response.Write(rs(i).value)
                else
                    response.Write("&nbsp;")
                end if
                %>
                </div>
            </td>
            <%
			next
			rs.movenext
		next 
        %>
        </tr>
    </table>
    <%end if%>        
</body>
</html>
<!--#include virtual="/Components/CopyRight.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->

->

