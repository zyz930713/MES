<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/RPT_DAILY_TARGET/RPT_DAILY_TARGETCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->

<%
fromdate=request("fromdate")
todate=request("todate")
if isnull(fromdate) or fromdate=""  then
	fromdate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if
todate=request("todate")
if isnull(todate) or todate=""  then	
	todate=cstr(year(now()))+"-"+cstr(month(now()))+"-"+cstr(day(now()))
end if


LINE=request("LINE")


'strWhere="where  Plan_Date BETWEEN TO_DATE('"&fromdate&"','yyyy-mm-dd') and TO_DATE('"&todate&"','yyyy-mm-dd') "

if LINE<>"" then
strWhere="where LINE like '%"&LINE&"%'"
end if



pagepara="&LINE="&LINE&"&fromdate="&fromdate&"&todate="&todate

SQL="select * from RPT_DAILY_TARGET "&strWhere&" order by LINE"

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
<form action="/Admin/RPT_DAILY_TARGET/RPT_DAILY_TARGET.asp" method="get" name="form1" target="_self">
               


<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td height="20"><span id="inner_Line"></span></td>
    <td height="20"><input name="LINE" type="text" id="LINE" value="<%=LINE%>"></td>
    
   
 </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()">
	
	</td>
    </tr>
</table>




</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="15" class="t-c-greenCopy"><span id="inner_BrowseData"></span></td>
</tr>
<tr>
  <td height="20" colspan="15" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="50%"><span id="inner_User"></span>:
          <% =session("User") %></td>
      <td width="50%"><div align="right"><%if admin=true then%>
      <a href="/Admin/RPT_DAILY_TARGET/AddRPT_DAILY_TARGET.asp?path=<%=path%>&query=<%=query%>" target="main" class="white"><span id="inner_AddRecord"></span></a>
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
  <td height="20" colspan="2" class="t-t-Borrow"><div align="center"><span id="inner_Action"></span></div></td>
  <%end if%>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_SubSeries"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_LineName"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_ProdcutTarget"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_ProcessYieldTarget"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_AcousticYieldTarget"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_FOIYieldTarget"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_LeakageTarget"></span></div></td>
  <td height="20" class="t-t-Borrow"><div align="center"><span id="td_ACOUSITC_OFFSET"></span></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="td_TTYTarget"></span></div></td>
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
  
  
  <div align="center"><span class="red"style="cursor:hand" onClick="javascript:window.open('EditRPT_DAILY_TARGET.asp?PRODUCT=<%=rs("PRODUCT")%>&Line=<%=rs("Line")%>&path=<%=path%>&query=<%=query%>','main')"><img src="/Images/IconEdit.gif" alt="Click to edit"></span></div>
 


  </td>
  <td height="20"><div align="center"><span class="red" style="cursor:hand" onClick="javascript:if (confirm('Are you sure to delete this section?')){window.open('EditRPT_DAILY_TARGET1.asp?PRODUCT=<%=rs("PRODUCT")%>&Line=<%=rs("Line")%>&path=<%=path%>&query=<%=query%>&Action=Del','main')}"><img src="/Images/IconDelete.gif" alt="Click to delete" width="16" height="20"></span></div></td>
    <%end if%>
    <td height="20"><div align="center"><%= rs("PRODUCT") %>&nbsp;</div></td>
    <td height="20"><div align="center"><%= rs("LINE") %>&nbsp;</div></td>
	<td><div align="center"><%= rs("TARGET")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("PROCESS_YIELD_TARGET")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("ACOUSTIC_YIELD_TARGET")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("FOI_YIELD_TARGET")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("LEAKAGE_YIELD_TARGET")%>&nbsp;</div></td>
    <td><div align="center"><%= rs("ACOUSITC_OFFSET")%>&nbsp;</div></td>
	<td><div align="center"><%= rs("FPY_TARGET")%>&nbsp;</div></td>
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