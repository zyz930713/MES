<%@  language="VBSCRIPT" codepage="936" %>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/PageSelect_new.asp" -->
<%
action=request("Action")
fromdate=request("fromdate")
jobnumber=request("jobnumber")
sheetNumber=request("sheet_number")
code=request("2d_code")
boxid=request("box_id")

if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
todate=request("todate")
if isnull(todate) or todate=""  then	
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if

if action="GenereateReport" then
	
	strWhere=" WHERE LM_TIME BETWEEN TO_DATE('"&fromdate&"','yyyy-mm-dd hh24:mi:ss') and TO_DATE('"&todate&" 23:59:59','yyyy-mm-dd hh24:mi:ss') "		
	if jobnumber<>"" then
		strWhere=strWhere+" and job_number='"+jobnumber+"'"
		if sheetNumber <> "" then
			strWhere=strWhere+" and sheet_number='"+sheetNumber+"'"
		end if
	end if 
	if code<>"" then
		strWhere=strWhere+" and code='"+code+"'"
	end if
	if boxid<>"" then
		strWhere=strWhere+" and box_id='"+boxid+"'"
	end if
	sql="SELECT CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,BOX_ID,B.DEFECT_CODE,B.DEFECT_NAME,B.DEFECT_CHINESE_NAME,rownum rn FROM JOB_2D_CODE A LEFT JOIN DEFECTCODE B ON A.DEFECT_CODE_ID = B.NID "	
	sql="select * from ("&sql&strWhere&" and rownum<="&pagenum*recordsize&") where rn>"&(pagenum-1)*recordsize
	session("totalSql")="select count(*) from job_2d_code a left join defectcode b on a.defect_code_id = b.nid "&strWhere
	rs.open sql,conn,1,3
end if
pagepara="&Action="&action&"&fromdate="&fromdate&"&todate="&todate&"&jobnumber="&jobnumber&"&sheet_number="&sheetNumber&"&2d_code="&code&"&box_id="&boxid
pagename="Job2DCode1.asp"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <title>Barcode System - Scan </title>
    <link href="/CSS/General.css" rel="stylesheet" type="text/css">
    <link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
	<script language="javascript" src="/Language/language.js" type="text/javascript"></script>
    <script language="javascript" src="/Components/sniffer.js" type="text/javascript"></script>
    <script language="javascript" src="/Components/dynCalendar.js" type="text/javascript"></script>
    <script type="text/javascript">
        function GenerateReport() {
            form1.action = "Job2DCode1.asp?Action=GenereateReport"
            form1.submit();
        }
		
    </script>
</head>
<body onLoad="language('<%=session("language")%>');language_page();">
    <form action="Job2DCode1.asp" method="post" name="form1" target="_self" id="form1">
    <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
        <tr>
            <td height="20" colspan="11" class="t-c-greenCopy"><span id="inner_2DCodeReport"></span></td>
        </tr>
        <tr align="center">
            <td width="80">
                <span id="inner_SearchJobNumber"></span>
            </td>
            <td width="200">
                <input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>" > - 
				<input name="sheet_number" type="text" id="sheet_number" size="4" value="<%=sheetNumber%>" >
            </td>
			<td width="80">
                <span id="inner_2DCode"></span>
            </td>
            <td width="100">
               <input name="2d_code" type="text" id="2d_code" value="<%=code%>" size="30" >
            </td>
            
            <td width="100"><span id="inner_BoxID"></span></td>
            <td width="100"><input name="box_id" type="text" id="box_id" value="<%=boxid%>"></td>
            <td width="100">
                <span id="inner_linkTime"></span>&nbsp;<span id="inner_SearchFrom"></span>
            </td>
            <td width="100">
                <input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
                <script language="JavaScript" type="text/javascript">
                    function calendar1Callback(date, month, year) {
                        document.all.fromdate.value = year + '-' + month + '-' + date
                    }
                    calendar1 = new dynCalendar('calendar1', 'calendar1Callback', document.all.fromdate.value);
                </script>               
            </td>
            <td width="30">
                <span id="inner_SearchTo"></span>
            </td>
            <td width="100">
                <input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
                <script language="JavaScript" type="text/javascript">
                    function calendar2Callback(date, month, year) {
                        document.all.todate.value = year + '-' + month + '-' + date
                    }
                    calendar2 = new dynCalendar('calendar2', 'calendar2Callback', document.all.todate.value);
                </script>
          <td align="left">&nbsp;
                    <img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle"
                        style="cursor: hand" onclick="GenerateReport()" >
          </td>
        </tr>
		<tr>
			<td height="20" colspan="18" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
		  </tr>
		  <tr>
			<td height="20" colspan="18" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>:
				  <% =session("User") %>        </td>					
			  </tr>
			</table></td>
		  </tr>		
    </table>
    </form>
    <%if(rs.State > 0 ) then %>
    <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF"> 
	  	<tr>
    		<td height="20" colspan="16">
			<!--#include virtual="/Components/PageSplit_new.asp" -->
			</td>
  		</tr>
        <tr>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_2DCode"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_JobNumber"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_SheetNumber"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_linkUser"></span></div></td>	
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_linkTime"></span></div></td>
			<td height="20" class="t-t-Borrow"><div align="center"><span id="td_BoxId"></span></div></td>
			<td height="20" class="t-t-Borrow" colspan="3"><div align="center"><span id="td_DefectCode"></span></div></td>              
        </tr>
        <%j=1
		while not rs.eof%>
        <tr>
			<td align="center"><%=(cint(pagenum)-1)*recordsize+j%></td>
            <%for i=0 to rs.Fields.count-2%>
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
			j=j+1
			rs.movenext
		wend 
        %>
        </tr>
</table>
    <%end if%>        
</body>
</html>
<!--#include virtual="/Components/CopyRight.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->

