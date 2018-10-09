<%
factorywhereoutside=""
factorywhereoutsidenull=""
factorywhereoutsideand=""
factorywhereoutsideandnull=""
factorywhereinside=""
factoryitem=""
function FactoryRight(prefix)
	user_factory_array=split(session("factory"),",")
	if instr(session("role"),",FACTORY_ADMINISTRATOR")=0 and instr(session("role"),",FACTORY_READER")=0 then
		factorywhereoutcore=" ("&prefix&"FACTORY_ID in ("		
		factorywhereincore=" ( NID in ("
		facSqlIn = ""
		for ui=0 to ubound(user_factory_array)
			facSqlIn = facSqlIn & " '" & user_factory_array(ui) & "' ,"		
		next
		if len(facSqlIn) > 0 then
			facSqlIn = left(facSqlIn,len(facSqlIn)-1)
		else
			facSqlIn = "''"
		end if
		factorywhereoutside=" where "&factorywhereoutcore & facSqlIn &")) "
		factorywhereoutsidenull=" where "&factorywhereoutcore & facSqlIn &") or "&prefix&"FACTORY_ID is null) "
		factorywhereoutsideand=" and "&factorywhereoutcore & facSqlIn &")) "
		factorywhereoutsideandnull=" and "&factorywhereoutcore & facSqlIn &") or "&prefix&"FACTORY_ID is null) "
		factorywhereinside=" where "&factorywhereincore & facSqlIn &")) "
		
		factoryitem=session("factory")
	end if
end function

modelfactory=""
modelfactoryand=""
function ModelFactoryRight()
	user_factory_array=split(session("factory"),",")
	if instr(session("role"),",FACTORY_ADMINISTRATOR")=0 and instr(session("role"),",FACTORY_READER")=0 then
		select case session("factory")
		case "FA00000001"
		modelfactory=" where (wip_supply_subinventory='C10 DELTEK'"
		modelfactoryand =" and segment1 like '____-______' and (wip_supply_subinventory='C10 DELTEK'"
		case "FA00000002"
		modelfactory=" where (wip_supply_subinventory='C10 ASSY'"
		modelfactoryand=" and (segment1 like '%-_____-%' or segment1 like 'SR%') and (wip_supply_subinventory='C10 ASSY'"
		case "FA00000003"
		modelfactory=" where (wip_supply_subinventory='C10 VALADD'"
		modelfactoryand =" and (wip_supply_subinventory='C10 VALADD'"
		end select
		modelfactory=modelfactory&" or wip_supply_subinventory is null or wip_supply_subinventory='') "
		modelfactoryand=modelfactoryand&" or wip_supply_subinventory is null or wip_supply_subinventory='') "
	end if
end function
%>