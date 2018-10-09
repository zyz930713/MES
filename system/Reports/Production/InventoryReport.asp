<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->

<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->

<%
set rssum=server.createobject("adodb.recordset")
Status=request("Status")

ProdType=request("ProdType")

if ProdType="" then
ProdType="FG"
end if

PART_NUMBER=trim(request("PART_NUMBER"))

plan_id=request("plan_id")
if plan_id<>"" then
strWhere="where plan_id='"&plan_id&"'"
orderby="Shipped_Time"
else
if Status="Inventory" then
strWhere=" WHERE WHREC_TIME is not null and whrec_time>'2013-08-23 23:59:59' and  SHIPPED_TIME is null and BOXIDSTATUS='"&ProdType&"'"

orderby="WHREC_TIME"
elseif Status="Geting" then

strWhere=" WHERE GET_TIME is not null and GET_time>'2013-08-23 23:59:59' and STACK_TIME is null and  SHIPPED_TIME is null and BOXIDSTATUS='"&ProdType&"'"

orderby="GET_TIME"


elseif  Status="Stacking" then

strWhere=" WHERE WHREC_TIME is not null and whrec_time>'2013-08-23 23:59:59' and STACK_TIME is not null and  SHIPPED_TIME is null and BOXIDSTATUS='"&ProdType&"'"

orderby="STACK_TIME"

else
strWhere=" WHERE WHREC_TIME is not null and whrec_time>'2013-08-23 23:59:59' and STACK_TIME is not null and  SHIPPED_TIME is not null and BOXIDSTATUS='"&ProdType&"'"

orderby="SHIPPED_TIME"
end if

end if

if PART_NUMBER<>"" then
strWhere=strWhere+"and PART_NUMBER='"&PART_NUMBER&"'"

pagepara="&PALLET_ID="&PALLET_ID&"&Status="&Status&"&ProdType="&ProdType&"&PART_NUMBER="&PART_NUMBER

else
pagepara="&PALLET_ID="&PALLET_ID&"&Status="&Status&"&ProdType="&ProdType&"&Plan_id="&Plan_id
end if

SQL="select box_id, part_number, SUM(packed_qty) as packed_qty, WHREC_USER,WHREC_Time,BOXIDSTATUS,GET_USER,GET_Time,PALLET_ID,stack_user, stack_time,shipped_user, shipped_time from JOB_PACK_DETAIL "&strWhere&" group by box_id,part_number,WHREC_USER,WHREC_Time ,GET_USER,GET_Time,PALLET_ID, stack_user, stack_time, BOXIDSTATUS,shipped_user, shipped_time order by "&orderby&" desc"
'response.Write(sql)
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

</head>

<body onLoad="language_page();language(<%=session("language")%>);">
<form action="/Reports/Production/InventoryReport.asp" method="get" name="form1" target="_self">
               


<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="6" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
  <td width="32%"><select name="ProdType"  id="ProdType">
    <option value="FG">良品</option>
    <option value="SCRAP">不良品</option>
    <option value="MRB">封存</option>
    </select></td>
   <td width="5%"  height="20"><span id="inner_SearchPartNumber"></span></td>
    <td width="10%" ><input name="PART_NUMBER" type="text" id="PART_NUMBER" value="<%=PART_NUMBER%>"></td>
    <td width="6%" height="20"><span id="inner_SearchStatus"></span></td>
    <td width="19%"><select name="Status"  id="Status">
                 
                  <option value="Inventory">库存</option>
                  <option value="Geting">出库</option>
				   <option value="Stacking">已堆拍</option>
				  <option value="Shipping">已发货</option>
				   
				   
                </select></td>
    <td width="28%"><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">	</td>
   </tr>
     
</table>


</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="14" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
  <td height="20" colspan="14"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td> 
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_BoxId"></span></div></td>

  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Model"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_Quantity"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_WHRECUser"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_WHRECTime"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_WHLOC"></span></div></td>
   <td height="20" class="t-t-Borrow"><div align="center"><span id="td_GetUser"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_GetTime"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_PalletID"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_StackUser"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_StackTime"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_ShippedUser"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_ShippedTime"></span></div></td>

  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr align="center">
  <td height="20"><div align="center">
    <% =(cint(session("strpagenum"))-1)*recordsize+i%>
  </div>  </td>
 
    <td height="20"><%= rs("BOX_ID") %></td>
	
	<td><%= rs("PART_NUMBER")%>&nbsp;</td>
	<td><%= rs("PACKED_QTY")%>&nbsp;</td>
    <td><%= rs("WHREC_USER")%></td>
    <td><%= rs("WHREC_TIME")%>&nbsp;</td>
    <td><%= rs("BOXIDSTATUS")%>&nbsp;</td>
    <td><%= rs("GET_USER")%>&nbsp;</td>
    <td><%= rs("GET_TIME")%>&nbsp;</td>
	<td><%= rs("PALLET_ID")%>&nbsp;</td>
	<td><%= rs("STACK_USER")%>&nbsp;</td>
	<td><%= rs("STACK_TIME")%>&nbsp;</td>
	<td><%= rs("SHIPPED_USER")%>&nbsp;</td>
	<td><%= rs("SHIPPED_TIME")%>&nbsp;</td>
	
   
  </tr>
<%
i=i+1
rs.movenext
wend
rs.close%>

<%


sqlsum="select sum(qtysum) as qtysum from (select  sum(packed_qty) as qtysum from JOB_PACK_DETAIl "&strWhere&" group by box_id)"

'response.Write(sqlsum)
rssum.open sqlsum,conn,1,3
if  not rssum.bof then
qtysum=rssum("qtysum")
end if
rssum.close%>
<tr>
    <td height="20" colspan="14" align="center"><span id="inner_Quantity"></span>:<%=qtysum%></td>
  </tr>
<%

else

%>
  <tr>
    <td height="20" colspan="14" align="center"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
'%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->