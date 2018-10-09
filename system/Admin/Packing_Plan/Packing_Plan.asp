<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Packing_Plan/Packing_PlanCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->

<%
fromdate=request("fromdate")
todate=request("todate")
plan_id=request("plan_id")
'if isnull(fromdate) or fromdate=""  then
	'fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
'end if
'todate=request("todate")
'if isnull(todate) or todate=""  then	
'	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
'end if


PARTNUMBER=request("PARTNUMBER")
NewStatus=request("NewStatus")
factory=request("factory")

if plan_id<>"" then
strWhere="where plan_id='"&plan_id&"'"
else



if PARTNUMBER<>"" then
strWhere="where PART_NUMBER like '%"&PARTNUMBER&"%'"
elseif  fromdate="" and  todate="" then
 if NewStatus<>"" then
 strWhere="where STATUS='"&NewStatus&"'"
 else
 strWhere="where STATUS<>'Complete'"
 end if
else


if fromdate<>"" and  todate <>"" and PARTNUMBER<>"" then
strWhere=strWhere&" and  Plan_Date BETWEEN TO_DATE('"&fromdate&"','yyyy-mm-dd') and TO_DATE('"&todate&"','yyyy-mm-dd') "
else
strWhere="where  Plan_Date BETWEEN TO_DATE('"&fromdate&"','yyyy-mm-dd') and TO_DATE('"&todate&"','yyyy-mm-dd') "
end if

end if
end if

pagepara="&PARTNUMBER="&PARTNUMBER&"&fromdate="&fromdate&"&todate="&todate&"&NewStatus="&NewStatus

SQL="select * from Packing_Plan "&strWhere&" order by DELIVERY_TIME"



rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charsPRet=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="javascript">
function formyieldcheck()
{
	with(document.formYield)
	{
		if(action.selectedIndex==0)
		{
			alert("Please select which action to be convert!")
			return false;
		}
	}
}
function queryDate(){
	if(!form1.PARTNUMBER.value){
		alert("Action Value cannot be blank. 数值不能为空。");
		return false;
	}
	document.form1.submit()	
}

</script>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form action="/Admin/Packing_Plan/Packing_Plan.asp" method="get" name="form1" target="_self">
               


<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_SearchPartNumber"></span></td>
    <td height="20"><input name="PARTNUMBER" type="text" id="PARTNUMBER" value="<%=PARTNUMBER%>"></td>
    <td><span id="inner_PlanDate"></span></td>
    <td><input name="fromdate" type="text" id="fromdate" value="" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><span id="inner_SearchStatus"></span>:</td>
    <td><select name="NewStatus">
     
    <option value="Run ">Run</option>
    <option value="Wait">Wait</option>
    <option value="Pending">Pending</option>
    <option value="Cancel">Cancel</option>
    <option value="Complete">Complete</option>
    </select></td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">
	
	</td>
    </tr>
</table>




</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
  <td height="20" class="t-c-greenCopy"><div align="right"><%if admin=true then%>
      <a href="/Admin/Plan_Customer_Name_Config/Plan_Customer_Name_Config.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_PlanCustomerNameConfig"></span></a>
    <%end if%></div></td>
</tr>
<tr>
  <td height="20" colspan="15" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/Packing_Plan/addPacking_Plan.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
      <%end if%></div></td>
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="15"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_PlanDate"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Model"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Quantity"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_StackQuantity"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_CustomerPN"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Remark"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Priority"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_DeliveryTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_CompleteTime"></span> </div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Status"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LMUser"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_LMTime"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div></td>
  <%if admin=true then%>
  <td>
  <%'if rs("STATUS")<>"Complete" then %>
  
  <div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditPacking_Plan.asp?PLAN_ID=<%=rs("PLAN_ID")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div>
  <%else%>
  &nbsp;
  <%'end if%>
  </td>
    <%end if%>
    <td height="20"><div align="center"><%= rs("PLAN_DATE") %></div></td>
	<td><a href="../../Reports/Production/InventoryReport.asp?Plan_id=<%=rs("Plan_id")%>"><%= rs("PART_NUMBER")%></a>&nbsp;</td>
	<td><%= rs("QUANTITY")%>&nbsp;</td>
    <td><%= rs("PACK_QTY")%>&nbsp;</td>
	<td><%= rs("CUSTOMER_NAME")%>&nbsp;</td>
	<td><%= rs("CUSTOMER_PART_NUMBER")%>&nbsp;</td>
	<td><%= rs("REMARK")%>&nbsp;</td>
    <td><div align="center"><%= rs("PRIORITY")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("DELIVERY_TIME")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("COMPLETE_TIME")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("STATUS")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("LM_USER")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("LM_TIME")%>&nbsp;</div></td>
</tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="15" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->