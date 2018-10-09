<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"
pagename="Station1.asp"
errorstring=request.QueryString("errorstring")
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>Job Tray Link</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript" src="../Scripts/jquery.blockUI.js"></script>
<script type="text/javascript" src="../Scripts/jquery.form.js"></script>
<script type="text/javascript" src="../Scripts/common.js"></script>
<script type="text/javascript" src="../Scripts/KES1/JobTrayLink.js"></script>
</head>
<body bgcolor="#339966">
<form id="form1">
  <table id="tb1" width="90%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
    <thead>
      <tr>
        <td colspan="6" class="t-t-DarkBlue">Job Tray Link 工单料盘绑定</td>
      </tr>
      <tr>
        <td>Sub Job Number 子工单号<span class="red"> *</span></td>
        <td><input type="text" name="txt_jobno" id="txt_jobno" /></td>
        <td>Operator Code 工号<span class="red"> *</span></td>
        <td><input type="text" name="txt_opcode" id="txt_opcode" /></td>
      </tr>
      <tr>
        <td>Start Qty 开始数量</td>
        <td><input type="text" name="txt_qty" id="txt_qty" disabled="disabled" /></td>
        <td>Part Number 料号</td>
        <td><input type="text" name="txt_partno" id="txt_partno" disabled="disabled" /></td>
      </tr>
    </thead>
    <tbody>
    </tbody>
    <tfoot>
      <tr> </tr>
    </tfoot>
  </table>
  <div id="div_station">
  </div>
</form>
</body>
</html>
