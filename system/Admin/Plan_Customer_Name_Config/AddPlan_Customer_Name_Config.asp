<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/Plan_Customer_Name_Config/Plan_Customer_Name_Config.aspCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
path=request.QueryString("path")
query=request.QueryString("query")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/Section/FormCheck.js" type="text/javascript"></script>
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<script language="javascript"  type="text/javascript">
	function CheckForm()
{
 
 
    if (document.form1.PART_NUMBER.value=="")
  {
    alert("12NC不能为空！");
	document.form1.PART_NUMBER.focus();
	return false;
   }else if (document.form1.PART_NUMBER.value.length!=12)
  {
    alert("12NC号位数不对！");
	document.form1.PART_NUMBER.focus();
	return false;
  
  }else if (document.form1.CUSTOMER_NAME.value=="")
  {
    alert("客户名称不能为空！");
	document.form1.CUSTOMER_NAME.focus();
	return false;
  }
  
  
 
}


 function checkNumber() {    //判断输入字符的keyCode，数字在48到57之间，超出部分返回false   
	  if ((event.keyCode >= 48) && (event.keyCode <= 57))
	   {      
	     event.returnValue = true;    
		 } 
		 else 
		 {        event.returnValue = false;    }
		 
		 }

</script>
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
<form action="/Admin/Plan_Customer_Name_Config/AddPlan_Customer_Name_Config1.asp" method="post" name="form1" target="_self" onSubmit="return CheckForm()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>


<tr>
  <td width="5%" height="20"><span id="inner_PartNumber"></span> <span class="red">*</span></td>
  <td width="95%" height="20"><input name="PART_NUMBER" type="text" id="PART_NUMBER"  value=""  onkeyup="lookup(this.value);" onblur="fill();" >
  <div class="suggestionsBox" id="suggestions" style="display: none;">
   <div class="suggestionList" id="autoSuggestionsList">
   &nbsp;
   </div>
   </div></td>
</tr>

<tr>
  <td height="20"><span id="td_CustomerName"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_NAME" type="text" id="CUSTOMER_NAME"  value=""></td>
</tr>





  <tr>
    <td height="20" colspan="2"><div align="center">
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