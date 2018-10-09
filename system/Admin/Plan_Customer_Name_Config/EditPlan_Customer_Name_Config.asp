<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Plan_Customer_Name_Config/Plan_Customer_Name_Config.aspCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
path=request.QueryString("path")
query=request.QueryString("query")
CONFIGID=request.QueryString("CONFIGID")
set rs1=server.CreateObject("adodb.recordset")
SQL="select *  from Plan_Customer_Name_Config  where CONFIGID='"&CONFIGID&"'"

rs.open SQL,conn,1,3
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Functions/JCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Functions/FromTo.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript">
 function lookup(inputString) {
  if(inputString.length == 0) {
   // Hide the suggestion box.
   $('#suggestions').hide();
  } else {
   $.post("showmember.asp", {queryString: ""+escape(inputString)+""}, function(data){
    if(data.length >0) {
     $('#suggestions').show();
     $('#autoSuggestionsList').html(unescape(data));
    }
   });
  }
 } // lookup
 
 function fill(thisValue) {
  $('#PART_NUMBER').val(thisValue);
  setTimeout("$('#suggestions').hide();", 200);
 }
</script>
<style type="text/css">
 body {
  font-family: Helvetica;
  font-size: 11px;
  color: #000;
 }
 
 h3 {
  margin: 0px;
  padding: 0px; 
 }

 .suggestionsBox {
  position: relative;
  left: 5px;
  margin: 10px 0px 0px 0px;
  width: 200px;
  background-color: #212427;
  -moz-border-radius: 7px;
  -webkit-border-radius: 7px;
  border: 2px solid #000; 
  color: #fff;
 }
 
 .suggestionList {
  margin: 0px;
  padding: 0px;
 }
 
 .suggestionList li {
  
  margin: 0px 0px 3px 0px;
  padding: 3px;
  cursor: pointer;
 }
 
 .suggestionList li:hover {
  background-color: #659CD8;
 }
</style>



</head>

<body onLoad="language(<%=session("language")%>);">
<form action="/Admin/Plan_Customer_Name_Config/EditPlan_Customer_Name_Config1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>


<tr>
  <td height="20"><span id="inner_PartNumber"></span> <span class="red">*</span></td>
  <td height="20"><input type="text" id="PART_NUMBER" value="<%=rs("PART_NUMBER")%>" name="PART_NUMBER"  onkeyup="lookup(this.value);" onblur="fill();">
  <div class="suggestionsBox" id="suggestions" style="display: none;">
   <div class="suggestionList" id="autoSuggestionsList">
   &nbsp;
   </div>
   </div> </td>
</tr>

<tr>
  <td height="20"><span id="td_CustomerName"></span> <span class="red">*</span></td>
  <td height="20"><input type="text" value="<%=rs("CUSTOMER_NAME")%>" name="CUSTOMER_NAME"></td>
</tr>







  <tr>
    <td height="20" colspan="2"><div align="center">
    <input name="CONFIGID" type="hidden" id="CONFIGID" value="<%=CONFIGID%>">
	
      <input name="path" type="hidden" id="path" value="<%=path%>">
      <input name="query" type="hidden" id="query" value="<%=query%>">
      <input type="submit" name="btnOK" value="OK">
&nbsp;
<input name="Reset" type="reset" id="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->