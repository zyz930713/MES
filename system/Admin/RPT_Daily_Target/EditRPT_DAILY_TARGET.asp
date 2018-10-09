<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/RPT_Daily_Target/RPT_DAILY_TARGETCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Components/FactoryRight.asp" -->

<%
path=request.QueryString("path")
query=request.QueryString("query")
PRODUCT=trim(request("PRODUCT"))
LINE=trim(request("LINE"))





set rs1=server.CreateObject("adodb.recordset")
SQL="select *  from RPT_Daily_Target  where product='"&PRODUCT&"' and Line='"&Line&"'"

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
    alert("泄漏良率目标不对！");
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
<form action="/Admin/RPT_Daily_Target/EditRPT_DAILY_TARGET1.asp" method="post" name="form1" target="_self" onSubmit="return CheckForm()">

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>

<tr>
  <td width="13%" height="20"><span id="td_SubSeries"></span> <span class="red">*</span></td>
  <td width="87%" height="20"> <input name="PRODUCT" type="text" id="PRODUCT"  value="<%=rs("PRODUCT")%>" readonly>
    
	  </td>
</tr>
<tr>
  <td height="20"><span id="td_LineName"></span> <span class="red">*</span></td>
  <td height="20"><input name="LINE" type="text" id="LINE"  value="<%=rs("LINE")%>"  readonly="" ></td>
</tr>
<tr>
  <td height="20"><span id="td_ProdcutTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="TARGET" type="text" id="TARGET"  value="<%=rs("TARGET")%>"  onkeypress="return checkNumber(event,this);"></td>
</tr>
<tr>
  <td height="20"><span id="td_ProcessYieldTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="PROCESS_YIELD_TARGET" type="text" id="PROCESS_YIELD_TARGET"  value="<%=rs("PROCESS_YIELD_TARGET")%>"  onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')" ></td>
</tr>
<tr>
  <td height="20"><span id="td_AcousticYieldTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="ACOUSTIC_YIELD_TARGET" type="text" id="ACOUSTIC_YIELD_TARGET"  value="<%=rs("ACOUSTIC_YIELD_TARGET")%>" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_FOIYieldTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="FOI_YIELD_TARGET" type="text" id="FOI_YIELD_TARGET"  value="<%=rs("FOI_YIELD_TARGET")%>" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_LeakageTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="LEAKAGE_YIELD_TARGET" type="text" id="LEAKAGE_YIELD_TARGET"  value="<%=rs("LEAKAGE_YIELD_TARGET")%>" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
</tr>
<tr>
  <td height="20"><span id="td_ACOUSITC_OFFSET"></span> <span class="red">*</span></td>
  <td height="20"><input name="ACOUSITC_OFFSET" type="text" id="ACOUSITC_OFFSET"  value="<%=rs("ACOUSITC_OFFSET")%>" ></td>
</tr>
<tr>
  <td height="20"><span id="td_TTYTarget"></span> <span class="red">*</span></td>
  <td height="20"><input name="FPY_TARGET" type="text" id="FPY_TARGET"  value="<%=rs("FPY_TARGET")%>" onblur="this.value=this.value.replace(/^(?!0\.(?!0{2})\d{2}$).+$/,'')"></td>
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
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->