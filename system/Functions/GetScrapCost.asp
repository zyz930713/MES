<%
Function GetScrapCost(nid)
	set objRst=Server.CreateObject("adodb.recordset")
	SQL="select JS.*, GET_ITEM_PRICE(JOB_NUMBER) as ITEM_COST from JOB_MASTER_SCRAP JS where NID='" & trim(nid&"") & "'"
	
	objRst.open SQL,conn,1,3
	if not objRst.eof then

		item_cost=objRst("ITEM_COST")
		total_qty=clng(objRst("SCRAP_QUANTITY"))
		'response.Write(objRst("PART_NUMBER_TAG") & "   " & objRst("SCRAP_QUANTITY") & "   " & objRst("ITEM_COST"))	' & "   " & objRst("CURRENCY"))
		GetScrapCost=cdbl(item_cost)*clng(total_qty)
	else
		GetScrapCost=0
	end if
	objRst.close
	set objRst=nothing
End Function

Function GetItemCostByJob(jobnumber)
	set objRst=Server.CreateObject("adodb.recordset")
	SQL="select GET_ITEM_PRICE('" & jobnumber & "') as ITEM_COST from dual"
	
	objRst.open SQL,conn,1,3
		
	
	if not objRst.eof then
		item_cost=objRst("ITEM_COST")
		if trim(item_cost&"")="" then
			item_cost=0
		end if
		GetItemCostByJob=cdbl(item_cost)
	else
		GetItemCostByJob=0
	end if
	objRst.close
	set objRst=nothing

End Function
%>