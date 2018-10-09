 <%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Series/SeriesCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->
<!--#include virtual="/Functions/GetFactory.asp" -->
<!--#include virtual="/Functions/GetModel.asp" -->
<%
model_name=request("model_name")
SQL="SELECT PM.ITEM_NAME,PM.DESCRIPTION FROM PRODUCT_MODEL PM WHERE PM.ITEM_STATUS='ACTIVE' and FAMILY_ID is null "
includesModel=request("models")
if len(includesModel)>1 then
	includesModel=left(includesModel,len(includesModel)-1)
	SQL=SQL & " AND ITEM_NAME NOT IN ('" & replace(includesModel,",","','") &"')"
end if
if(not isnull(request("model_name")) and request("model_name")<>"") then
	SQL=SQL & " AND ITEM_NAME LIKE '" & model_name &"%'"
end if
pagename="/Admin/SubSeries_New/ModelSelect.asp"
pagepara="&model_name="&model_name

'response.write SQL
session("SQL")=SQL
rs.open SQL,conn,1,3
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<base target="_self"/>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<script type="text/javascript">
	function Return_Model()
	{
		var PageCount;
		PageCount=parseFloat(document.getElementById("pagecount").value);
		var ModelList;
		ModelList="";
		for(var i=1;i<PageCount;i++)
		{
			
			if (document.getElementById("id"+i).checked==true)
			{
				ModelList=ModelList+document.getElementById("txtModelName"+i).innerHTML+",";
			}
		}
		
		dialogArguments.document.getElementById("txtIncludeModel").value=dialogArguments.document.getElementById("txtIncludeModel").value+ModelList;
		window.close();
	}
	function SelectAll()
	{
		var count=document.getElementById("pagecount").value;
		var selAll=document.getElementById("sel_all").checked;
		for(var i=1;i<count;i++)
		{
			document.getElementById("id"+i).checked=selAll;
		}
	}
</script>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<!--#include virtual="/Language/Components/Lan_PageSplit.asp" -->
</head>
<body onLoad="language_page();language(<%=session("language")%>);">
<form action="ModelSelect.asp" method="post" name="form1" >
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td colspan="3" class="t-b-midautumn"><span id="inner_Search"></span></td>
  </tr>
  <tr>
    <td width="80" ><span id="inner_SearchPartNumber"></span></td>
    <td width="100"><input name="model_name" type="text" id="model_name" value="<%=model_name%>"></td>
    <td align="left" ><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>
</form>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="13"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td class="t-t-Borrow"><div align="center"><input type="checkbox" name="sel_all" id="sel_all" onClick="SelectAll()"></div></td>
  <td class="t-t-Borrow"><div align="center"><span id="inner_NO"></span></div></td>  
  <td class="t-t-Borrow"><div align="center"><span id="td_Model"></span></div></td>
  <td class="t-t-Borrow"><span id="td_Description"></span></td>
  </tr>
<form name="checkform" method="post">
<%
i=1
if not rs.eof then
rs.absolutepage=cint(session("strpagenum"))
while not rs.eof and i<=rs.pagesize 
%>
<tr>
	 <td align="center"><input name="id<%=i%>" type="checkbox" id="id<%=i%>" value="1"></td>
  	<td align="center"><%=(cint(session("strpagenum"))-1)*pagesize_s+i%></td>	
    <td height="20"><div align="center"  name="txtModelName<%=i%>" id="txtModelName<%=i%>"><%= rs("ITEM_NAME") %></div></td>
    <td><%= rs("DESCRIPTION") %>&nbsp;</td>
  </tr>
<%
i=i+1
rs.movenext
wend
%>
  
<%
else
%>
  <tr>
    <td height="20" colspan="13"><span id="inner_Records"></span>&nbsp;</td>
  </tr>
<%end if
rs.close%>

<Tr>
	<td colspan="4">
		&nbsp; <input name="pagecount" type="hidden" id="pagecount" value="<%=i%>">
		<input name="btnOK" type="submit" id="btnOK" onClick="Return_Model()" value="OK">
	</td>
</Tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
