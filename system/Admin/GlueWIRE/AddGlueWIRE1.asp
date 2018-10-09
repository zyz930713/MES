<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/Admin/PACKING/PACKINGCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
ITEM_NAME=trim(request("ITEM_NAME"))
SUPPLIER_NAME=trim(request("SUPPLIER_NAME"))
STATION_DESCRIPTION=trim(request("STATION_DESCRIPTION"))
STATION_DESCRIPTION_EN=trim(request("STATION_DESCRIPTION_EN"))
set rs1=server.createobject("adodb.recordset")
SQL="select msi.segment1 as ITEM_NAME from apps.mtl_system_items@keprod msi inner join apps.cst_item_costs@keprod cic on msi.inventory_item_id=cic.inventory_item_id AND msi.organization_id = cic.organization_id where msi.organization_id = 495 and msi.segment1='"&ITEM_NAME&"' AND cic.cost_type_id = 1"
rs1.open SQL,conn,1,3
if rs1.bof and  rs1.eof then

word="12NC"&TEST_NAME&"不正确"
'action="location.href='"&beforepath&"'"
action="history.back()"

else

function repeatstring(str,ex,num)
	repeatstring=string(num-len(str),ex)&str
end function

countType="GlueWIRE"
countCondition="SN"
sql="select count_value,lm_time,count_type,count_condition from serial_count "
sql=sql+"where count_type='"&countType&"' and count_condition = '"&countCondition&"'"


rs.open sql,conn,1,3
if rs.eof then
STATION_NO=countCondition&"0001"
rs.addNew
rs("count_type")=countType
rs("count_condition")=countCondition
rs("count_value")=1
rs("lm_time")=now()
else
rs("count_value")=clng(rs("count_value"))+1
STATION_NO=countCondition&repeatstring(rs("count_value"),"0",4)
rs("lm_time")=now()
end if
rs.update
rs.close		
 
sql="select ITEM_NAME from Product_model where ITEM_NAME='"&ITEM_NAME&"'"
rs.open sql,conn,1,3
if rs.bof and rs.eof then
conn.execute("insert into PRODUCT_MODEL (ITEM_ID,FACTORY_ID,ITEM_NAME,DESCRIPTION,ITEM_TYPE,LAST_UPDATE_DATE,CURRENCY,ITEM_COST,ITEM_STATUS) select cic.inventory_item_id as ITEM_ID, 'FA00000028' as FACTORY,msi.segment1 as ITEM_NAME, msi.description as DESCRIPTION,msi.item_type as ITME_TYPE,msi.last_update_date LAST_UPDATE_DATE, 'CNY' as CURRENCY, cic.item_cost ITEM_COST, msi.inventory_item_status_code as ITEM_STATUS from apps.mtl_system_items@keprod msi inner join apps.cst_item_costs@keprod cic on msi.inventory_item_id=cic.inventory_item_id AND msi.organization_id = cic.organization_id where msi.organization_id = 495 and msi.segment1='"&ITEM_NAME&"' AND cic.cost_type_id = 1")
conn.Execute("update Product_model  set SUPPLIER_NAME='"&SUPPLIER_NAME&"' where ITEM_NAME='"&ITEM_NAME&"'")
else
conn.Execute("update Product_model  set SUPPLIER_NAME='"&SUPPLIER_NAME&"' where ITEM_NAME='"&ITEM_NAME&"' ")
end if 
rs.close
 
SQL="select * from MATERIAL_STATION where ITEM_NAME='"&ITEM_NAME&"' and  STATION_DESCRIPTION='"&STATION_DESCRIPTION&"'  and  STATION_DESCRIPTION_EN='"&STATION_DESCRIPTION_EN&"'"
rs.open SQL,conn,1,3
if rs.eof then
rs.addnew
rs("STATION_NO")=STATION_NO
rs("ITEM_NAME")=ITEM_NAME
rs("STATION_DESCRIPTION")=STATION_DESCRIPTION
rs("STATION_DESCRIPTION_EN")=STATION_DESCRIPTION_EN
rs.update
rs.close
word="Successfully save a New GlueWire config."
action="location.href='"&beforepath&"'"
else
word="12NC、站点、项目已存在！"
action="history.back()"
end if

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