<%
function get_single_quantity_amount(part_number,quantity)
set rsS=server.CreateObject("adodb.recordset")
SQLS="select nvl(S1.ITEM_COST,0) as SINGLE1_COST,nvl(S1.ITEM_COST,0) as SINGLE2_COST from DUAL_SETTINGS DS left join PRODUCT_MODEL S1 on DS.SINGLE1=S1.ITEM_NAME left join PRODUCT_MODEL S2 on DS.SINGLE2=S2.ITEM_NAME where DS.DUAL_NAME='"&part_number&"'"
rsS.open SQLS,conn,1,3
if not rsS.eof then
	get_single_quantity_amount=(ccur(rsS("SINGLE1_COST"))+ccur(rsS("SINGLE2_COST")))*csng(quantity)
else
	get_single_quantity_amount=0
end if
rsS.close
set rsS=nothing
end function
%>