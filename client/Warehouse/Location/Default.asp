<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>功能选择</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
</head>
<body style="margin:20px;background-color:#339966;">
<center>
<table width="500" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF" >
	<tr>
		<td height="20" colspan="2" class="t-t-DarkBlue" align="center">
			Select Station 选择所需操作
		</td>
	</tr>
	<tr>
		<td colspan = "2">　</td>
	</tr>
	<tr>
		<td align="center">
			入出库操作
		</td>
		<td>
			<input type="button" class="t-b-midautumn" value="入库" onClick="javascript:window.open('WarehouseIn.asp')">
			&nbsp;
			<input type="button" class="t-b-midautumn" value="出库" onClick="javascript:window.open('WarehouseOut.asp')">
            &nbsp;
			<input type="button" class="t-b-midautumn" value="移库" onClick="javascript:window.open('WarehouseMove.asp')">
		</td>
	</tr>
	<tr>
		<td colspan = "2">　</td>
	</tr>
	<tr>
		<td align="center">
			记录查询
		</td>
		<td>
			<input type="button" class="t-b-midautumn" value="入出库记录" onClick="javascript:window.open('WarehouseRec.asp')">
			&nbsp;
			<input type="button" class="t-b-midautumn" value="库存信息" onClick="javascript:window.open('WarehouseItem.asp')">
		</td>
	</tr>
    <tr>
		<td colspan = "2">　</td>
	</tr>
	<tr>
		<td align="center">系统设置</td>
		<td>
			<input type="button" class="t-b-midautumn" value="同步12NC" onClick="javascript:window.open('Set_Sync12NC.asp')">
			&nbsp;
			<input type="button" class="t-b-midautumn" value="供应商绑定" onClick="javascript:window.open('Set_Item_Supplier.asp')">
		</td>
	</tr>
</table>
</center>
</body> 
</html>