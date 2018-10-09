<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
action=request("action")
StartTime=request("txtfromdate")
if StartTime="" then
	StartTime=date() 
	firstTime=dateAdd("d",-5,date())
else
	firstTime=dateAdd("d",-5,StartTime)
end if

if action<>"1" then
 addwhere=" and 1<>1 "
end if
SQL="select job_number,sub_joblist,transaction_type,(select actualqty From label_print_history where job_number=aa.job_number and subjoblist=aa.sub_joblist and decode(transactiontype,'-1','N','1','R','5','S')=aa.TRANSACTION_TYPE and rownum=1 ) as qty,( select part_number_tag From job_master where job_number=aa.job_number) as part_number from ((select to_char(b.job_number) as job_number ,to_char(b.subjoblist) as sub_joblist ,decode(transactiontype,'-1','N','1','R','5','S') AS TRANSACTION_TYPE From job_retest a, label_print_history b where a.batchno=b.batchno "&addwhere&" and a.reteststarttime>=to_date('"&firstTime&" 00:00:00','yyyy-mm-dd hh24:mi:ss') and a.reteststarttime<=to_date('"&StartTime&" 23:59:59','yyyy-mm-dd hh24:mi:ss') ) MINUS ( select to_char(job_number) as job_number,to_char(replace(replace(replace(replace(sub_job_numbers,job_number,''),'-00',''),'-0','') ,',','-')) as subjoblist,store_type AS TRANSACTION_TYPE From job_master_store a )) aa order by aa.job_number"

rs.open SQL,conn,1,3
session("SQL")=SQL
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Untitled Document</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
	
</head>

<body>
<form name="form1" method="post" action="/Job/Job/UnInstoreJob.asp?action=1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Search">Search Jobs</span></td>
  </tr>
  <tr>
    
    
    <td><span id="inner_SearchStartTime">Date</span></td>
    <td>
      <input name="txtfromdate" type="text" id="txtfromdate" value="<%=StartTime%>" size="10">
      <script language=JavaScript type=text/javascript>
		function calendar1Callback(date, month, year)
		{
		document.all.txtfromdate.value=year + '-' + month + '-' + date
		}
		calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
      </script>
		&nbsp;<span id="inner_SearchEndTime"></span>
	 </td>
	  <td ><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
	 
    </tr>
 
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><span id="inner_Browse">Job List</span></td>
</tr>
<tr>
  <td height="20" colspan="12" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User">User:</span>: 
          <% =session("User") %></td>
         <td width="50%"><div align="right">
          <span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('UnInstoreJob_Export.asp')"><img src="/Images/EXCEL_Middle.gif" width="22" height="22"></span>&nbsp;</span></div></td>
    </table></td>
</tr>
 <tr>
      <td height="20" colspan="12"></td>
    </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO">No.</span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_JobNumber">Job Number</span></div></td> 
  <td class="t-t-Borrow"><div align="center">Sub Job List</div></td>
  <td class="t-t-Borrow"><div align="center">Transaction Type</div></td>
  <td class="t-t-Borrow"><div align="center">Quantity</div></td>
  <td class="t-t-Borrow"><div align="center">Part Number</div></td>
 </tr>
<%
	i=1
if not rs.eof then
'rs.absolutepage=cint(session("strpagenum"))
while not rs.eof 
%>
<Tr>
  <td height="20"><div align="center"><%=i%></div></td>  
  <td height="20"><div align="center"><%=rs("job_number")%></div></td>
  <td ><div align="center"><%=rs("sub_joblist")%></div></td>  
  <td ><div align="center"><%=rs("transaction_type")%>&nbsp;</div></td>
  <td ><div align="center"><%=rs("qty")%>&nbsp;</div></td>
  <td ><div align="center"><%=rs("part_number")%>&nbsp;</div></td>
 </tr>
 
<%	
	i=i+1
	rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="12"><div align="center">No Records </div></td>
  </tr>
<%end if
rs.close%>
 
</table>
<!--#include virtual="/Components/CopyRight.asp" -->

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->