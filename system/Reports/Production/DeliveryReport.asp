<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
if instr(session("role"),",DELIVERY_ADMINISTRATOR")<>0 then
	admin=true
end if
fromdate=request("fromdate")
todate=request("todate")
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
todate=request("todate")
if isnull(todate) or todate=""  then	
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if

deliveryId=request("txt_deliveryId")

strWhere=" where shipped_time between to_date('"&fromdate&"','yyyy-mm-dd hh24:mi:ss') and to_date('"&todate&" 23:59:59','yyyy-mm-dd hh24:mi:ss') "
if deliveryId<>"" then
	strWhere=" where delivery_id like '%"&deliveryId&"%'"
else
	strWhere=strWhere&" and delivery_id is not null"
end if

pagepara="&txt_deliveryId="&deliveryId&"&fromdate="&fromdate&"&todate="&todate

SQL="select delivery_id,pallet_id,shipped_user,shipped_time,vendor_receive_date,confirm_user,plan_id,sum(packed_qty) as qty,confirm_remarks from job_pack_detail "&strWhere&" group by delivery_id,pallet_id,shipped_user,shipped_time,vendor_receive_date,confirm_user,plan_id,confirm_remarks order by vendor_receive_date desc"

rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
	function editDelivery(deliveryId){
		window.showModalDialog("EditDelivery.asp?id="+deliveryId,window,"dialogHeight:400px;dialogWidth:700px");		
		location.href="DeliveryReport.asp?1=1<%=pagepara%>";
	}
</script>
</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form method="post" name="form1" target="_self">  
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_DeliveryID"></span></td>
    <td height="20"><input name="txt_deliveryId" type="text" id="txt_deliveryId" value="<%=deliveryId%>"></td>    
    <td><span id="inner_ShippedTime"></span></td>
    <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;<span id="inner_SearchTo"></span>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">	</td>
    </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="11" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>      
    </tr>
  </table></td>
</tr>
<tr>
  <td height="20" colspan="11"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>
  <%if admin=true then%>
  <td class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>  
  <td class="t-t-Borrow"><div align="center"><span id="td_DeliveryID"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_PalletID"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Quantity"></span></div></td>  
  <td class="t-t-Borrow"><div align="center"><span id="td_ShippedUser"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_ShippedTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_VendorReceiveDate"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_ConfirmUser"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_PlanId"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_Remark"></span></div></td>
</tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr align="center">
  <td height="20"><div align="center"><% =(cint(session("strpagenum"))-1)*recordsize+i%></div></td>
  <%if admin then%>
  <td> 
  	<%if isnull(rs("vendor_receive_date")) then%>	
  		<span class="red"style="cursor:hand" onClick="editDelivery('<%=rs("delivery_id")%>')">
			<img src="/Images/IconEdit.gif" alt="Click to edit">
		</span>
	<%end if%>
	&nbsp;</td>
    <%end if%> 
    <td height="20"><div align="center"><%= rs("delivery_id") %></div></td>
	<td><%= rs("pallet_id")%>&nbsp;</td>
	<td><%= rs("qty")%>&nbsp;</td>
	<td><%= rs("shipped_user")%>&nbsp;</td>
	<td><%= rs("shipped_time")%>&nbsp;</td>
	<td><%= rs("vendor_receive_date")%>&nbsp;</td>
	<td><%= rs("confirm_user")%>&nbsp;</td>
	<td><a href="/admin/packing_plan/Packing_Plan.asp?plan_id=<%= rs("plan_id")%>" ><%= rs("plan_id")%>&nbsp;</a></td>
	<td><%= rs("confirm_remarks")%>&nbsp;</td>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="11" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->