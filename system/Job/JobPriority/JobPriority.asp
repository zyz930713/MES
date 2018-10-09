<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
 
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Job/JobCheck.asp" -->
<!--#include virtual="/Job/IsDBA.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
JobNumber=request("txtJobNumber")
SubJobNumber=request("txtSubJobNumber")
PRIORITY_LEVEL=request("Priority")
action=request.QueryString("Action")
 
if(action="1") then
	SQL="SELECT JOB_NUMBER,SHEET_NUMBER,JOB_PRIORITY FROM JOB WHERE JOB_NUMBER='"+JobNumber+"' AND SHEET_NUMBER='"+SubJobNumber+"'"
	set rsJobPriority=Server.CreateObject("adodb.recordset")
	rsJobPriority.open SQL,conn,1,3
	if rsJobPriority.recordcount>0 then
		if not rsJobPriority.eof then
			jobNo=rsJobPriority("JOB_NUMBER")
			sheetNo=rsJobPriority("SHEET_NUMBER")
			PRIORITY_LEVEL=rsJobPriority("JOB_PRIORITY")
		end if 
	end if 
end if 

 
if(action="2") then
	SQL="UPDATE JOB SET JOB_PRIORITY='"+PRIORITY_LEVEL+"' WHERE JOB_NUMBER='"+request("txtJobNo")+"' AND SHEET_NUMBER='"+request("txtSheetNo")+"'"
	set rsJobPriorityUpdate=Server.CreateObject("adodb.recordset")
	rsJobPriorityUpdate.open SQL,conn,1,3
	
	SQL="UPDATE TBL_MES_LOTMASTER SET JOB_PRIORITY='"+PRIORITY_LEVEL+"' WHERE wipentityname='"+request("txtJobNo")+"'"
	set rsJobPriorityUpdate_Ticket=Server.CreateObject("adodb.recordset")
	rsJobPriorityUpdate_Ticket.open SQL,connTicket,1,3
		
	response.write "<script>alert('Update Successfully!更新成功!')</script>"
end if 

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script>
	function SearchJobPriority()
	{
		if(document.getElementById("txtJobNumber").value=="")
		{
			alert("Please input job number!\n请输入工单号!");
			return;
		}
		if(document.getElementById("txtSubJobNumber").value=="")
		{
			alert("Please input Sheet number!\n请输入分批号!");
			return;
		}
		form1.action="JobPriority.asp?Action=1";
		form1.submit();
	}
	
	function SaveData()
	{		
		form1.action="JobPriority.asp?Action=2";
		form1.submit();
	}
</script>
</head>

<body onLoad="language(<%=session("language")%>);" >
<form action="JobPriority.asp" method="post" name="form1" target="_self">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
    <td height="20" colspan="5"  class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
    <tr>
      <td height="20" width="100"><span id="inner_SearchJobNumber"></span></td>
	   <td height="20"  width="50"><input type="text" id="txtJobNumber" name="txtJobNumber" ></td>
       <td height="20" width="100" ><span id="inner_SheetNumber"></span></td>
	   <td height="20"  width="50"><input type="text" id="txtSubJobNumber" name="txtSubJobNumber" ></td>
	   
	   <td height="20" align="left"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onclick="SearchJobPriority()">
	   </td>
    </tr>
	<table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#666666" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
  </tr>
  <tr>
    <td height="20" colspan="16" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%" class="t-c-greenCopy"><span id="inner_User"></span>:
          <% =session("User") %>
		</td>        
      </tr>
    </table></td>
  </tr>
	  <tr class="t-t-Borrow" align="center">
	  	<td height="20"><span id="td_JobNumber"></span> </td>
		<td height="20"><span id="td_SheetNumber"></span> </td>
      	<td height="20"><span id="td_JobPriority"></span> </td>
	  </tr>
	  <%if jobNo <> "" then%>
	  <tr align="center">
	  	<td><%=jobNo%></td>
		<td><%=sheetNo%></td>
	   	<td height="20">
	   	<%
			SQL="SELECT PRIORITY_LEVEL,PRIORITY_DEC FROM JOB_PRIORITY_SETTING order by PRIORITY_LEVEL"
			set rsPriority=Server.CreateObject("adodb.recordset")
			rsPriority.open SQL,conn,1,3
		%>
	   		<select name="Priority"  id="Priority" >				  
				  <% while not rsPriority.eof%>
				  	<option value=<%=rsPriority("PRIORITY_LEVEL")%> <%if PRIORITY_LEVEL=rsPriority("PRIORITY_LEVEL") then response.write " selected " end if %>><%=rsPriority("PRIORITY_DEC")%></option>
				  <%
				  		rsPriority.movenext
				  	wend
					rsPriority.close
				  %>
		</select>
	   </td>
    </tr>
	  <tr>
      <td height="20" colspan="3">
	  	<input type="hidden" id="txtJobNo" name="txtJobNo" value="<%=jobNo%>" >
	  	<input type="hidden" id="txtSheetNo" name="txtSheetNo" value="<%=sheetNo%>" >		
	  	<input type="button" id="btnOK" name="btnOK" value="OK" onClick="SaveData()">
	  </td>
    </tr>
	<%end if%>	
</table>
<br>
  <!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_TICKET_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
