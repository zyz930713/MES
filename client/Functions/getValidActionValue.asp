<%
function getValidActionValue(job_number,part_number_id,action,max_quantity,min_quantity)
getValidActionValue=""
ValidActionValue=""
set rsA=server.CreateObject("adodb.recordset")
SQL="select * from ACTION where NID='"&action&"'"
'session("aerror")=SQL
rsA.open SQL,conn,1,3
if not rsA.eof then
	select case rsA("ACTION_PURPOSE")
	case "1" 'for machine
		has_valid_machine=false
		set rsV=server.CreateObject("adodb.recordset")
		' Job in Group
		SQLV="select MACHINE from GROUP_ACTION_VALUE where GROUP_ID=(select NID from SYSTEM_GROUP where STATUS=1 and GROUP_TYPE='JOB' and GROUP_MEMBERS like '%"&job_number&"%') and ACTION_ID='"&action&"'"
		'session("aerror")=SQLV
		rsV.open SQLV,conn,1,3
		if not rsV.eof then
			if not isnull(rsV("MACHINE")) and rsV("MACHINE")<>"" then
				has_valid_machine=true
				ValidActionValue=rsV("MACHINE")
			end if
		end if
		rsV.close
		' Part has machine value
		if has_valid_machine=false then
			SQLV="select MACHINE from PART_ACTION_VALUE where PART_ID='"&part_number_id&"' and ACTION_ID='"&action&"'"
			rsV.open SQLV,conn,1,3
			if not rsV.eof then
				if not isnull(rsV("MACHINE")) and rsV("MACHINE")<>"" then
					has_valid_machine=true
					ValidActionValue=rsV("MACHINE")
				end if
			end if
			rsV.close
			set rsV=nothing
		end if
		' default machine
		if not isnull(rsA("VALID_MACHINE")) and has_valid_machine=false then
			ValidActionValue=rsA("VALID_MACHINE")
		end if
	case "2" 'for material part number or material lot number
		set rsV=server.CreateObject("adodb.recordset")
		SQLV="select MATERIAL from PART_ACTION_VALUE where PART_ID='"&part_number_id&"' and ACTION_ID='"&action&"'"
		rsV.open SQLV,conn,1,3
		if not rsV.eof then
			if not isnull(rsV("MATERIAL")) then
				ValidActionValue=rsV("MATERIAL")
			end if
		end if
		rsV.close
		set rsV=nothing
	case "3"
		set rsV=server.CreateObject("adodb.recordset")
		SQLV="select MATERIAL from PART_ACTION_VALUE where PART_ID='"&part_number_id&"' and ACTION_ID='"&action&"'"
		rsV.open SQLV,conn,1,3
		if not rsV.eof then
			if not isnull(rsV("MATERIAL")) then
				ValidActionValue=rsV("MATERIAL")
			end if
		end if
		rsV.close
		set rsV=nothing
	case "4" 'for material quantity
		if rsA("MAX_QUANTITY")<>"0" then
			max_quantity=cint(rsA("MAX_QUANTITY"))
		end if
		if rsA("MIN_QUANTITY")<>"0" then
			min_quantity=cint(rsA("MIN_QUANTITY"))
		end if
	end select
end if
rsA.close
set rsA=nothing
getValidActionValue=ValidActionValue
end function
%>
