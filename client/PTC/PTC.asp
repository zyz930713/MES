<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
	response.Expires=0
	response.CacheControl="no-cache"
	pagename="Station.asp"
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.STYLE4 {font-size: 12px}
-->
</style>
</head>


<body bgcolor="#339966">
<form action="Station1.asp" method="post" name="form1">
<table width="461" height="48" align="center">
  <tr>
    <td align="center">PTC - Product Tracking &amp; Control (产品追溯和控制)</td>
  </tr></table>
<table border="0" align="center"><tr><td><table>
    <tr>
      <td height="20" ><div align="center">
	     <input name="btnPrintLable" type="button" class="t-b-midautumn" id="btnPrintLable" onClick="javascript:window.open('2DSan.asp')" value="2D码查询">       &nbsp;&nbsp;
	      <input name="btnPrintLable" type="button" class="t-b-midautumn" id="btnPrintLable" onClick="javascript:window.open('SanProd.asp')" value="单号查询">
		  &nbsp;&nbsp;
          <input name="btnPrintLable" type="button" class="t-b-midautumn" id="btnPrintLable" onClick="javascript:window.open('SendProd.asp')" value="借出">
          &nbsp;
          <input name="btnOQC" type="button" class="t-b-midautumn" id="btnOQC" onClick="javascript:window.open('Acceptprod.asp')" value="接收，归还">
          &nbsp;
          <input name="btnPreStore" type="button" class="t-b-midautumn" id="btnPreStore" onClick="javascript:window.open('EndProd.asp')" value="NPI 接收">
		  &nbsp;</div></td>
    </tr>
    <tr>
      <td height="20" ><div align="center">&nbsp;&nbsp;</div></td>
    </tr>
  </table>
</td></tr></table>
<input type="hidden" id="computername" name="computername">
</form>	
<table width="91%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="50%" height="20"><span class="STYLE4">&copy;  2013 - 2013 KEB Production System (BPS) 1.0.0</span></td>
    <td width="50%" height="20"><div align="right" class="STYLE4">Knowles Electronics (Beijing) </div></td>
  </tr>
  <tr>
    <td width="50%" height="20"><span class="STYLE4">Administrator: <a href="mailto:Young.Li@knowles.com">Li, Shaoyong (Young)</a>&nbsp;</span> </td>
    <td width="50%" height="20"><div align="right" class="STYLE4"> IT客服热线(1891 073) <B>6955</B> </div></td>
  </tr>
</table>


</body>

<script>
	var wsh=new ActiveXObject("WScript.Network"); 
	document.form1.computername.value=wsh.ComputerName; 
</script>
</html>

<!--#include virtual="/WOCF/BOCF_Close.asp" -->