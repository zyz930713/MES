<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%session("Page_Role") = ",WH_Location_Set"%>
<!--#include virtual="/Functions/Login_Check.asp"-->
<!--#include virtual="/Functions/Role_Check.asp"-->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/include/Functions.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>同步12NC</title>
<link href="Styles/Basic.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript">
</script>
<style type="text/css">
table tr{
	background:#cccccc;
}
tr.change:hover
{
	background-color:#66A6FF;
}
</style>
</head>
<body style="margin:20px;background-color:#339966;">
<%
btSearch = request("btSearch")
ItemName = trim(request("ItemName"))

if btSearch = "拷贝" or btSearch = "更新" then
	itemID = request("itemID")
	factoryID = "FA00000028"
	itemNameOr = request("itemNameOr")
	description = request("description")
	itemType = request("itemType")
	lastUpdatDate = request("lastUpdatDate")
	currencyType = "CNY"
	itemCost = request("itemCost")
	itemStatus = request("itemStatus")
	unitType = request("unitType")

	if btSearch = "拷贝" then
		' itemID2 = request("itemID")
		' itemNameOr2 = request("itemNameOr")
		' description2 = request("description")
		' itemType2 = request("itemType")
		' lastUpdatDate2 = request("lastUpdatDate")
		' itemCost2 = request("itemCost")
		' itemStatus2 = request("itemStatus")
		' unitType2 = request("unitType")
		exeSql = "insert into PRODUCT_MODEL(ITEM_ID,FACTORY_ID,ITEM_NAME,DESCRIPTION,ITEM_TYPE,LAST_UPDATE_DATE,CURRENCY,ITEM_COST,ITEM_STATUS,UNIT)"
		exeSql = exeSql & " values('"&itemID&"','"&factoryID&"','"&itemNameOr&"','"&description&"','"&itemType&"','"&lastUpdatDate&"','"&currencyType&"','"&itemCost&"','"&itemStatus&"','"&unitType&"')"
	elseif btSearch = "更新" then
		exeSql = "update PRODUCT_MODEL set DESCRIPTION = '"&description&"',UNIT = '"&unitType&"' where ITEM_NAME = '"&itemNameOr&"'"
	end if
	'response.write exeSql
	conn.execute(exeSql)
	response.write "<script>alert('同步成功');</script>"
	btSearch = "ok"
	ItemName = itemNameOr
end if



%>
<form name="form1" method="post" action="Set_Sync12NC.asp" >
<table width="1202" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr><td height="20" class="t-t-DarkBlue" align="center">同步12NC</td></tr>	
	<tr>
		<td align="center">料号：<input type="text" name="ItemName" value="<%=ItemName%>" />&nbsp;<input type="submit" name="btSearch" value="查询" class="t-b-midautumn" /></td>
	</tr>
</table>
</form>
<%
if btSearch = "查询" or btSearch = "ok" then
	dim countSum,countSyn
	set rs = Server.CreateObject("adodb.recordset")
	set rs2 = Server.CreateObject("adodb.recordset")
	sql = "select cic.inventory_item_id as ITEM_ID,msi.segment1 as ITEM_NAME,msi.description as DESCRIPTION,msi.item_type as ITEM_TYPE,msi.last_update_date LAST_UPDATE_DATE,'CNY' as CURRENCY,cic.item_cost ITEM_COST,msi.inventory_item_status_code as ITEM_STATUS,msi.primary_uom_code from apps.mtl_system_items@keprod msi inner join apps.cst_item_costs@keprod cic on msi.inventory_item_id=cic.inventory_item_id AND msi.organization_id = cic.organization_id where msi.organization_id = 495 and cic.cost_type_id = 1 and msi.segment1 = '"&ItemName&"'"
	'response.write sql
	rs.open sql,conn,1,1
	if not rs.eof then
		itemID = rs("ITEM_ID")
		itemNameOr = rs("ITEM_NAME")
		description = rs("DESCRIPTION")
		itemType = rs("ITEM_TYPE")
		lastUpdatDate = rs("LAST_UPDATE_DATE")
		'currencyType = rs("CURRENCY")
		itemCost = rs("ITEM_COST")
		itemStatus = rs("ITEM_STATUS")
		unitType = rs("primary_uom_code")
	%>
	<form name="form2" method="post" action="Set_Sync12NC.asp" >
	<table width="1202" border="1" align="center" cellpadding="0" cellspacing="0" >
	<tr>
		<td align="center" colspan="3" class="t-t-DarkBlue">料号：<%=itemNameOr%>
			<input type="hidden" name="itemID" value="<%=itemID%>"/>
			<input type="hidden" name="itemNameOr" value="<%=itemNameOr%>"/>
			<input type="hidden" name="description" value="<%=description%>"/>
			<input type="hidden" name="itemType" value="<%=itemType%>"/>
			<input type="hidden" name="lastUpdatDate" value="<%=lastUpdatDate%>"/>
			<input type="hidden" name="itemCost" value="<%=itemCost%>"/>
			<input type="hidden" name="itemStatus" value="<%=itemStatus%>"/>
			<input type="hidden" name="unitType" value="<%=unitType%>"/>
		</td>
	</tr>
	<tr><th width="50">&nbsp;</th><th>描述</th><th>单位</th></tr>
	<tr><td width="50">Remote</td><td><%=description%></td><td><%=unitType%></td></tr>
	<%
		sql2 = "SELECT ITEM_ID,FACTORY_ID,ITEM_NAME,DESCRIPTION,ITEM_TYPE,LAST_UPDATE_DATE,CURRENCY,ITEM_COST,ITEM_STATUS,nvl(UNIT,'NaN') UNIT FROM PRODUCT_MODEL where ITEM_NAME = '"&itemName&"'"
		rs2.open sql2,conn,1,1
		if rs2.eof then	'如果没有,那么就直接插入新的
			response.write "<td colspan='3' align='center'><input type='submit' name='btSearch' value='拷贝'/></td>"
		else	'如果系统中有,就检查单位是否为空
			response.write "<td colspan='3' align='center'><input type='submit' name='btSearch' value='更新'/></td>"
			itemID2 = rs2("ITEM_ID")
			itemNameOr2 = rs2("ITEM_NAME")
			description2 = rs2("DESCRIPTION")
			itemType2 = rs2("ITEM_TYPE")
			lastUpdatDate2 = rs2("LAST_UPDATE_DATE")
			'currencyType2 = rs2("CURRENCY")
			itemCost2 = rs2("ITEM_COST")
			itemStatus2 = rs2("ITEM_STATUS")
			unitType2 = rs2("UNIT")
			%>
			<tr><td width="50">Local</td><td><%=description2%></td><td><%=unitType2%>
					<input type="hidden" name="itemID2" value="<%=itemID2%>"/>
					<input type="hidden" name="itemNameOr2" value="<%=itemNameOr2%>"/>
					<input type="hidden" name="description2" value="<%=description2%>"/>
					<input type="hidden" name="lastUpdatDate2" value="<%=lastUpdatDate2%>"/>
					<input type="hidden" name="itemCost2" value="<%=itemCost2%>"/>
					<input type="hidden" name="itemStatus2" value="<%=itemStatus2%>"/>
					<input type="hidden" name="unitType2" value="<%=unitType2%>"/>
					</td></tr>
			<%
		end if
		rs2.close
%>
</table>
</form>
<%
	end if
end if
%>
<center>
  <table width="1202" border="1" cellspacing="0" cellpadding="0">
    <tr>
      <td align="center">
      <input type="button" value="修改密码" onClick="window.location.reload('../../Functions/UserPass_Modify.asp');" />
      　　　　
      <input type="button" value="Close关闭" onClick="window.location.reload('../../Functions/User_Logout.asp');" /></td>
    </tr>
  </table>
</center>
</body> 
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->