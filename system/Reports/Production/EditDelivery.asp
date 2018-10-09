<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->

<%
deliveryId=request("id")
isSubmit=request.Form("is_submit")
if isSubmit="1" then
	rcvDate=request.Form("txt_rcvDate")
	remarks=request.Form("txt_remarks")
	sql="update job_pack_detail set vendor_receive_date=to_date('"&rcvDate&"','yyyy-mm-dd hh24:mi:ss'),confirm_user='"&session("code")&"',confirm_remarks='"&remarks&"' where delivery_id='"&deliveryId&"'"
	conn.execute(sql)
	word="Successfully edit."
end if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Language/language.js" type="text/javascript"></script>
<script language="javascript" type="text/javascript" src="/Components/My97DatePicker/WdatePicker.js"></script>
<script type="text/javascript">
	function formcheck(){
		if(!form1.txt_rcvDate.value){
			alert("Vendor Receive Date cannot be blank.\n客户接收时间不能为空");
			return false;
		}
		return true;
	}
	if("<%=word%>" !="" ){
		alert("<%=word%>");
		window.close();
	}
</script>
</head>
<base target="_self">
<body onLoad="language(<%=session("language")%>);">
<form action="" method="post" name="form1" onSubmit="return formcheck()">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_EditData"></span> </td>
</tr>
	<tr> 
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_User"></span>:
        <% =session("User") %></td>
    </tr>

<tr>
  <td height="20"><span id="td_DeliveryID"></span> <span class="red">*</span></td>
  <td height="20"><%=deliveryId%></td>
</tr>
<tr>
  <td height="20"><span id="td_VendorReceiveDate"></span> <span class="red">*</span></td>
  <td height="20"><input name="txt_rcvDate" type="text" id="txt_rcvDate"   readonly="Ture" >
<img onclick="WdatePicker({el:'txt_rcvDate',dateFmt:'yyyy-MM-dd HH:mm:ss'})" src="/Images/dynCalendar.gif" align="absmiddle" style="cursor:pointer"/>
</td>
</tr>
<tr>
  <td height="20"><span id="td_Remark"></span></td>
  <td height="20"><input name="txt_remarks" type="text" id="txt_remarks" size="80"></td>
</tr>
  <tr>
    <td height="20" colspan="2"><div align="center">
	<input name="is_submit" type="hidden" id="is_submit" value="1">
	<input name="id" type="hidden" id="id" value="<%=deliveryId%>">
    <input type="submit" name="btnOK" value="OK">
	&nbsp;
	<input name="Reset" type="reset" id="btnReset" value="Reset">
</div></td>
    </tr>
</table>
</form>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->