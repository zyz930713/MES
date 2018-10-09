<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Admin/TrayLink/TrayLinkCheck.asp" -->
<!--#include virtual="/Admin/IsAdmin.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStation.asp" -->
<!--#include virtual="/Language/Admin/TrayLink/Lan_AddTrayLink.asp" -->
<%
path=request.QueryString("path")
query=request.QueryString("query")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<script language="JavaScript" src="/Admin/TrayLink/FormCheck.js" type="text/javascript"></script>
</head>
<body onLoad="language();">
<form action="/Admin/TrayLink/AddTrayLink1.asp" method="post" name="form1" target="_self" onSubmit="return formcheck()">
  <table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <tr>
      <td height="20" colspan="2" class="t-c-greenCopy"><span id="inner_NewTrayLink"></span></td>
    </tr>
    <tr>
      <td width="20%" height="20"><span id="inner_PartNumber"></span><span class="red">*</span> </td>
      <td width="80%" height="20" class="red"><input name="partnumber" type="text" id="partnumber" size="30"></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_StationName"></span><span class="red">*</span> </td>
      <td height="20" class="red"><div align="left">
          <select  name="stationname" id="stationname" >
		  <option value=""></option>
	  	<%=getStation_New(true,"OPTION","","","order by station_name","","") %>
		  </select>
		  	  	
        </div></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_StationSequence"></span><span class="red">*</span></td>
      <td height="20"><input name="stationseq" type="text" id="stationseq" size="30"></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_TrayType"></span><span class="red">*</span></td>
      <td height="20"><select name="traytype" id="traytype">
          <option value=""></option>
          <option value="OUT">OUT</option>
		  <option value="IN">IN</option>
        </select></td>
    </tr>
    <tr>
      <td height="20"><span id="inner_TraySize"></span><span class="red">*</span></td>
      <td height="20"><input name="traysize" id="traysize" size="30"/>
      </td>
    </tr>
    <tr>
      <td height="20" colspan="2"><div align="center">
          <input name="actionscount" type="hidden" id="actionscount" value="0">
          <input name="path" type="hidden" id="path" value="<%=path%>">
          <input name="query" type="hidden" id="query" value="<%=query%>">
          <input type="submit" name="btnOK" value="OK">
          &nbsp;
          <input type="reset" name="btnReset" value="Reset">
        </div></td>
    </tr>
  </table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
