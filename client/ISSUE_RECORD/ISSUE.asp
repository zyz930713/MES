<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">

</head>
<body onLoad="form1.txtSubJobList.focus();" bgcolor="#339966">

<table border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">  	
    <tr>
      <td height="20" class="t-t-DarkBlue"  align="center">(异常操作处理)</td>
    </tr>
	<%if request.QueryString("word") <> "" then%>
	<%end if%>
   <tr align="center"><td><input type="button" class="t-b-midautumn" value="解除单个二维码"onClick="javascript:window.open('/ISSUE_RECORD/Split_Code.asp')" >&nbsp;&nbsp;<input type="button" class="t-b-midautumn" value="正常整箱解箱号"onClick="javascript:window.open('/ISSUE_RECORD/Split_BoxAll.asp')" >&nbsp;
     <input type="button" class="t-b-midautumn" value="客退解箱号"onClick="javascript:window.open('/ISSUE_RECORD/Split_CodeRMA.asp')" >&nbsp;
     <input type="button" class="t-b-midautumn" value="临时库验证箱号"onClick="javascript:window.open('/ISSUE_RECORD/WarehouseRecTemp.asp')" >&nbsp;&nbsp; <input type="button" class="t-b-midautumn" value="临时库删除箱号"onClick="javascript:window.open('/ISSUE_RECORD/Split_Temp_Box.asp')" ></td></tr>
  </form>
</table>
</body>

</html>