<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/PACKING/PACKINGCheck.asp" -->

<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
STATION_NO=trim(request("STATION_NO"))
ITEM_NAME=trim(request("ITEM_NAME"))
SUPPLIER_NAME=trim(request("SUPPLIER_NAME"))
STATION_DESCRIPTION=trim(request("STATION_DESCRIPTION"))
STATION_DESCRIPTION_EN=trim(request("STATION_DESCRIPTION_EN"))
Action=trim(request("Action"))


if Action="Del" then



conn.Execute("Delete from MATERIAL_STATION  where STATION_NO='"&STATION_NO&"'")
word="Successfully Del Glue&Wire Config."
action="history.back()"


else
conn.Execute("update MATERIAL_STATION  set STATION_DESCRIPTION='"&STATION_DESCRIPTION&"',STATION_DESCRIPTION_EN='"&STATION_DESCRIPTION_EN&"' ,ITEM_NAME='"&ITEM_NAME&"' where STATION_NO='"&STATION_NO&"' ")

sql="select ITEM_NAME from Product_model where ITEM_NAME='"&ITEM_NAME&"'"
rs.open sql,conn,1,3
if rs.bof and rs.eof then
conn.execute("insert into PRODUCT_MODEL (ITEM_ID,FACTORY_ID,ITEM_NAME,DESCRIPTION,ITEM_TYPE,LAST_UPDATE_DATE,CURRENCY,ITEM_COST,ITEM_STATUS) select cic.inventory_item_id as ITEM_ID, 'FA00000028' as FACTORY,msi.segment1 as ITEM_NAME, msi.description as DESCRIPTION,msi.item_type as ITME_TYPE,msi.last_update_date LAST_UPDATE_DATE, 'CNY' as CURRENCY, cic.item_cost ITEM_COST, msi.inventory_item_status_code as ITEM_STATUS from apps.mtl_system_items@keprod msi inner join apps.cst_item_costs@keprod cic on msi.inventory_item_id=cic.inventory_item_id AND msi.organization_id = cic.organization_id where msi.organization_id = 495 and msi.segment1='"&ITEM_NAME&"' AND cic.cost_type_id = 1")
conn.Execute("update Product_model  set SUPPLIER_NAME='"&SUPPLIER_NAME&"' where ITEM_NAME='"&ITEM_NAME&"'")
else
conn.Execute("update Product_model  set SUPPLIER_NAME='"&SUPPLIER_NAME&"' where ITEM_NAME='"&ITEM_NAME&"' ")
end if
word="Successfully edit Glue&Wire Config."
action="location.href='"&beforepath&"'"

end if



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>
<script language="javascript">
alert("<%=word%>");
<%=action%>;
</script>
<body>

</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->