<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/PACKING/PACKINGCheck.asp" -->
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
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language="javascript" type="text/javascript" src="../../Components/My97DatePicker/WdatePicker.js"></script>
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language="javascript"  type="text/javascript">
	function CheckForm()
{
 
 
    if (document.form1.PRODUCT.value=="")
  {
    alert("子系列名称！");
	document.form1.PRODUCT.focus();
	return false;
   }else   if (document.form1.LINE.value=="")
  {
    alert("线别！");
	document.form1.LINE.focus();
	return false;
  }else   if (document.form1.TARGET.value=="")
  {
    alert("生产目标！");
	document.form1.TARGET.focus();
	return false;
  }else if (document.form1.PROCESS_YIELD_TARGET.value=="")
  {
    alert("流程良率目标！");
	document.form1.PROCESS_YIELD_TARGET.focus();
	return false;
  }
  
  else if (document.form1.ACOUSTIC_YIELD_TARGET.value=="")
  {
    alert("声学测试良率目标！");
	document.form1.ACOUSTIC_YIELD_TARGET.focus();
	return false;
  }
  
  else if (document.form1.FOI_YIELD_TARGET.value=="")
  {
    alert("外观检查良率目标！");
	document.form1.FOI_YIELD_TARGET.focus();
	return false;
  }
  
 else if (document.form1.LEAKAGE_YIELD_TARGET.value=="")
  {
    alert("泄漏良率目标！");
	document.form1.LEAKAGE_YIELD_TARGET.focus();
	return false;
  }
     else if (document.form1.ACOUSITC_OFFSET.value=="")
  {
    alert("声学补偿不对！");
	document.form1.ACOUSITC_OFFSET.focus();
	return false;
  }
  else if (document.form1.FPY_TARGET.value=="")
  {
    alert("工单总良率目标！");
	document.form1.FPY_TARGET.focus();
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

</head>

<body onLoad="language(<%=session("language")%>);">
<form action="/Admin/RPT_DAILY_TARGET/AddRPT_DAILY_TARGET1.asp" method="post" name="form1" target="_self" onSubmit="return CheckForm()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_AddData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>

<tr>
  <td width="13%" height="20"><span id="td_SubSeries"></span> <span class="red">*</span></td>
  <td width="87%" height="20"> <input name="PRODUCT" type="text" id="PRODUCT" >
    
	  </td>
</tr>
<tr>
  <td height="20"><span id="td_LineName"></span> <span class="red">*</span></td>
  <td height="20"><input name="LINE" type="text" id="LINE"  value=""  ></td>
</tr>
<tr>
  <td height="20"><span id="td_ProdcutTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="TARGET" type="text" id="TARGET"  value="" onkeypress="return checkNumber(event,this);" ></td>
</tr>
<tr>
  <td height="20"><span id="td_ProcessYieldTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="PROCESS_YIELD_TARGET" type="text" id="PROCESS_YIELD_TARGET"  value="" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_AcousticYieldTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="ACOUSTIC_YIELD_TARGET" type="text" id="ACOUSTIC_YIELD_TARGET"  value="" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_FOIYieldTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="FOI_YIELD_TARGET" type="text" id="FOI_YIELD_TARGET"  value="" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_LeakageTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="LEAKAGE_YIELD_TARGET" type="text" id="LEAKAGE_YIELD_TARGET"  value="" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_ACOUSITC_OFFSET"></span> <span class="red">*</span></td>
  <td height="20"><input name="ACOUSITC_OFFSET" type="text" id="ACOUSITC_OFFSET"  value="" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_TTYTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="FPY_TARGET" type="text" id="FPY_TARGET"  value="" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
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