<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>Sync Oracle 12NC</title>
</head>
<body>
<%
set connOra = Server.CreateObject("ADODB.Connection")
connOra.Open "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome4;User ID=BAR_WEB_KEB;Data Source=KEBDB"

dim countSum,countSyn
set rs = Server.CreateObject("adodb.recordset")
set rs2 = Server.CreateObject("adodb.recordset")
sql = "select cic.inventory_item_id as ITEM_ID,msi.segment1 as ITEM_NAME,msi.description as DESCRIPTION,msi.item_type as ITME_TYPE,msi.last_update_date LAST_UPDATE_DATE,'CNY' as CURRENCY,cic.item_cost ITEM_COST,msi.inventory_item_status_code as ITEM_STATUS,msi.primary_uom_code from apps.mtl_system_items@keprod msi inner join apps.cst_item_costs@keprod cic on msi.inventory_item_id=cic.inventory_item_id AND msi.organization_id = cic.organization_id where msi.organization_id = 495 and cic.cost_type_id = 1"
'response.write sql
rs.open sql,connOra,1,1

countSum = 0
countInsert = 0
countUpdate = 0
while not rs.eof
	itemID = rs("ITEM_ID")
	itemName = rs("ITEM_NAME")
	description = rs("DESCRIPTION")
	itemType = rs("ITME_TYPE")
	lastUpdatDate = rs("LAST_UPDATE_DATE")
	currencyType = rs("CURRENCY")
	itemCost = rs("ITEM_COST")
	itemStatus = rs("ITEM_STATUS")
	unitType = rs("primary_uom_code")

	sql2 = "SELECT ITEM_ID,FACTORY_ID,ITEM_NAME,DESCRIPTION,ITEM_TYPE,LAST_UPDATE_DATE,CURRENCY,ITEM_COST,ITEM_STATUS,nvl(UNIT,'NaN') UNIT FROM PRODUCT_MODEL where ITEM_NAME = '"&itemName&"'"
	rs2.open sql2,connOra,1,1
	if not rs2.eof then	'如果系统中有,就检查单位是否为空
		'if rs2("UNIT") = "" then
		'if isnull(rs2("UNIT")) then
		'if (description <> rs2("DESCRIPTION")) or (itemType <> rs2("ITEM_TYPE")) or (itemCost <> rs2("ITEM_COST")) or (unitType <> rs2("UNIT")) then
		if unitType <> rs2("UNIT") then
			exeSql = "update PRODUCT_MODEL set UNIT = '"&unitType&"' where ITEM_NAME = '"&itemName&"'"
			countUpdate = countUpdate + 1
		end if
		'response.write exeSql & "<br>"
		connOra.execute(exeSql)
	else	'如果没有,那么就直接插入新的
		'exeSql = "insert into PRODUCT_MODEL(ITEM_ID,FACTORY_ID,ITEM_NAME,DESCRIPTION,ITEM_TYPE,LAST_UPDATE_DATE,CURRENCY,ITEM_COST,ITEM_STATUS,UNIT)"
		'exeSql = exeSql & " values("&itemID&","&itemName&","&description&","&itemType&","&lastUpdatDate&","&currencyType&","&itemCost&","&itemStatus&","&unitType&")"
		countInsert = countInsert + 1
	end if
	rs2.close

	countSum = countSum + 1
	rs.movenext
wend
response.write "合计:" & countSum & "<br>"
response.write "更新:" & countUpdate & "<br>"
response.write "忽略:" & countInsert & "<br>"
%>
</body>
</html>