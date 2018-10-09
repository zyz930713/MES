<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/Packing_Plan/Packing_PlanCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
path=request.QueryString("path")
query=request.QueryString("query")
PLAN_ID=request.QueryString("PLAN_ID")
set rs1=server.CreateObject("adodb.recordset")
SQL="select *  from Packing_plan  where PLAN_ID='"&PLAN_ID&"'"
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
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language="javascript" type="text/javascript" src="../../Components/My97DatePicker/WdatePicker.js"></script>
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<script language="javascript"  type="text/javascript">
	function CheckForm()
{
 
 
    if (document.form1.PLAN_DATE.value=="")
  {
    alert("计划日期不能为空！");
	document.form1.PLAN_DATE.focus();
	return false;
   }else   if (document.form1.PART_NUMBER.value=="")
  {
    alert("料号不能为空！");
	document.form1.PART_NUMBER.focus();
	return false;
  }else   if (document.form1.QUANTITY.value=="")
  {
    alert("数量不能为空！");
	document.form1.QUANTITY.focus();
	return false;
  }else if (document.form1.CUSTOMER_NAME.value=="")
  {
    alert("客户名称不能为空！");
	document.form1.CUSTOMER_NAME.focus();
	return false;
  } 
  
  else if (document.form1.REMARK.value=="")
  {
    alert("备注不能为空！");
	document.form1.REMARK.focus();
	return false;
  }
  else if (document.form1.DELIVERY_TIME.value=="")
  {
    alert("交货时间不能为空！");
	document.form1.DELIVERY_TIME.focus();
	return false;
  }
   
  
   if (document.form1.NewStatus.value=="Cancel")
   
   {
	   if ( document.form1.StackQuantity.value>0)
         {
			alert("已有产品堆拍，请联系大库房处理！");	
			return false;
	
	      }

	   
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
 
 
 
  function lookup1(inputString) {
	 
  if(inputString.length == 0) {
   // Hide the suggestion box.
   $('#suggestions1').hide();
  } else {
   $.post("showmember.asp", {cstname: ""+escape(inputString)+""}, function(data){
    if(data.length >0) {
     $('#suggestions1').show();
     $('#autoSuggestionsList1').html(unescape(data));
    }
   });
  }
 } // lookup
 
 function fille(thisValue) {
  $('#CUSTOMER_NAME').val(thisValue);
  setTimeout("$('#suggestions1').hide();", 200);
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
<form action="/Admin/Packing_Plan/EditPacking_Plan1.asp" method="post" name="form1" target="_self" onSubmit="return CheckForm()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>

<tr>
  <td height="20"><span id="td_PlanDate"></span> <span class="red">*</span></td>
  <td height="20"><input type="text" id="PLAN_DATE" name="PLAN_DATE" value="<%=rs("PLAN_DATE")%>"> <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.PLAN_DATE.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
  </td>
</tr>
<tr>
  <td height="20"><span id="inner_PartNumber"></span> <span class="red">*</span></td>
  <td height="20"><input name="PART_NUMBER" type="text" id="PART_NUMBER"  value="<%=rs("PART_NUMBER")%>"  onkeyup="lookup(this.value);" onblur="fill();" >
   <div class="suggestionsBox" id="suggestions" style="display: none;">
    
    <div class="suggestionList" id="autoSuggestionsList">
     &nbsp;
    </div>
   </div></td>
</tr>
<tr>
  <td height="20"><span id="td_Quantity"></span> <span class="red">*</span></td>
  <td height="20"><input type="text" name="QUANTITY" value="<%=rs("QUANTITY")%>"></td>
</tr>
<tr>
  <td height="20"><span id="td_StackQuantity"></span> <span class="red">*</span></td>
  <td height="20"><input type="text" name="StackQuantity" value="<%=rs("PACK_QTY")%>" readonly style="background-color:#666666"></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerName"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_NAME" type="text" id="CUSTOMER_NAME" onblur="fille();" onkeyup="lookup1(document.getElementById('PART_NUMBER').value);"  value="<%=rs("CUSTOMER_NAME")%>" size="30">
   <div class="suggestionsBox" id="suggestions1" style="display: none;">
    
    <div class="suggestionList" id="autoSuggestionsList1">
     &nbsp;
    </div>
   </div></td>
</tr>
<tr>
  <td height="20"><span id="td_CustomerPN"></span> <span class="red">*</span></td>
  <td height="20"><input name="CUSTOMER_PART_NUMBER" type="text" id="CUSTOMER_PART_NUMBER"  value="<%=rs("CUSTOMER_PART_NUMBER")%>"></td>
</tr>
<tr>
  <td height="20"><span id="td_Remark"></span> <span class="red">*</span></td>
  <td height="20"><input name="REMARK" type="text" value="<%=rs("REMARK")%>" size="50"></td>
</tr>
<tr>
  <td height="20"><span id="td_Priority"></span> <span class="red">*</span></td>
  <td height="20"><%PRIORITY=rs("PRIORITY")%>
  
    <select name="PRIORITY" id="PRIORITY">
    
    <option value="1"<%if PRIORITY="1" then%>selected<%end if%>>1</option>
    <option value="2"<%if PRIORITY="2" then%>selected<%end if%>>2</option>
    <option value="3"<%if PRIORITY="3" then%>selected<%end if%>>3</option>
    <option value="4"<%if PRIORITY="4" then%>selected<%end if%>>4</option>
    <option value="5"<%if PRIORITY="5" then%>selected<%end if%>>5</option>
  </select>&nbsp;<span class="red" id="inner_PriorityDescription"></span></td>
</tr>

<tr>
  <td height="20"><span id="td_DeliveryTime"></span> <span class="red">*</span></td>
  <td height="20"><input name="DELIVERY_TIME" type="text" id="DELIVERY_TIME"   readonly="Ture" value="<%=rs("DELIVERY_TIME")%>" size="18">
<img onclick="WdatePicker({el:'DELIVERY_TIME',dateFmt:'yyyy-MM-dd HH:mm:ss',onpicked:pickedFunc})" src="../../Images/dynCalendar.gif" align="absmiddle" style="cursor:pointer"/>
<script>
function pickedFunc(){
 $dp.$('d523_y').value=$dp.cal.getP('y');
 $dp.$('d523_M').value=$dp.cal.getP('M');
 $dp.$('d523_d').value=$dp.cal.getP('d');
 $dp.$('d523_HH').value=$dp.cal.getP('H');
 $dp.$('d523_mm').value=$dp.cal.getP('m');
 $dp.$('d523_ss').value=$dp.cal.getP('s');
 }
</script>
</td>
</tr>
<tr>
  <td height="20"><span id="td_Status"></span> <span class="red">*</span></td>
  <td height="20"><%=rs("STATUS")%></td>
</tr>
<tr>
  <td height="20"><span id="td_NewStatus"></span> <span class="red">*</span></td>
  <td height="20"> <select name="NewStatus" id="NewStatus">
    
    <option value="Wait">Wait</option>
    <option value="Pending">Pending</option>
    <option value="Cancel">Cancel</option>
    <option value="Complete">Complete</option>
    
  </select></td>
</tr>

  <tr>
    <td height="20" colspan="2"><div align="center">
	<input name="PLAN_ID" type="hidden" id="PLAN_ID" value="<%=PLAN_ID%>">
	
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