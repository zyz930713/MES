<%
function NID_SEQ(table)
set rsE=server.CreateObject("adodb.recordset")
	select case table
	case "FACTORY" 'FA
		SQLE="select lpad(FACTORY_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "SECTION" 'SE
		SQLE="select lpad(SECTION_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "SERIES" 'SR
		SQLE="select lpad(SERIES_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "SERIES_GROUP" 'SG
		SQLE="select lpad(SERIES_GROUP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FINANCE_SERIES" 'FS
		SQLE="select lpad(FINANCE_SERIES_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FINANCE_SERIES_GROUP" 'FG
		SQLE="select lpad(FINANCE_SERIES_GROUP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "DUAL_SETTINGS" 'DU
		SQLE="select lpad(DUAL_SETTINGS_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "YIELD_EXCLUSION" 'YE
		SQLE="select lpad(YIELD_EXCLUSION_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "LINE" 'LI
		SQLE="select lpad(LINE_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "PART" 'PA
		SQLE="select lpad(PART_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "GUID" 'GU
		SQLE="select lpad(GUID_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "ROUTING" 'RT
		SQLE="select lpad(ROUTING_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "STATION" 'ST
		SQLE="select lpad(STATION_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "STATION_New" 'SA
		SQLE="select lpad(STATION_NEW_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "ACTION" ' AC
		SQLE="select lpad(ACTION_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "ACTION_New" ' AN
		SQLE="select lpad(ACTION_NEW_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "Routing_Action" ' RA
		SQLE="select lpad(Routing_Action_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "Routing_Defect" ' RD
		SQLE="select lpad(Routing_Defect_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "DEFECTCODE" 'DE
		SQLE="select lpad(DEFECTCODE_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "DEFECTCODE_New" 'DE
		SQLE="select lpad(DEFECTCODE_New_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "DEFECTCODE_GROUP" 'DG
		SQLE="select lpad(DEFECTCODE_GROUP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "MACHINE" 'MA
		SQLE="select lpad(MACHINE_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "MATERIAL"' MT
		SQLE="select lpad(MATERIAL_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "OPERATOR" 'OP
		SQLE="select lpad(OPERATOR_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "USER" 'US
		SQLE="select lpad(USER_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "USER_GROUP" 'UG
		SQLE="select lpad(USER_GROUP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "ROLE" 'RO
		SQLE="select lpad(ROLE_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "APPROVAL_ROLE" 'RO
		SQLE="select lpad(APPROVAL_ROLE_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "EVENT" 'EV
		SQLE="select lpad(EVENT_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "TASK" 'TA
		SQLE="select lpad(TASK_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FORM" 'FO
		SQLE="select lpad(FORM_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "PROFILETASK" 'PT
		SQLE="select lpad(PROFILE_TASK_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "PROFILEFORM" 'PF
		SQLE="select lpad(PROFILE_FORM_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "GROUP"  'GR
		SQLE="select lpad(GROUP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "DAILY_LINEYIELD" 'DY
		SQLE="select lpad(DAILY_LINEYIELD_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "DAILY_STOCKYIELD" 'SY
		SQLE="select lpad(DAILY_LINEYIELD_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FINAL_FAMILYYIELD" 'FF
		SQLE="select lpad(FINAL_FAMILYYIELD_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FINAL_SERIESYIELD" 'FS
		SQLE="select lpad(FINAL_SERIESYIELD_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FINAL_PARTYIELD" 'FP
		SQLE="select lpad(FINAL_PARTYIELD_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FAMILY_LINELOST" 'FL
		SQLE="select lpad(FAMILY_LINELOST_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FAMILY_SCRAP" 'FC
		SQLE="select lpad(FAMILY_SCRAP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "WIP" 'WI
		SQLE="select lpad(WIP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "PART_WIP" 'PW
		SQLE="select lpad(WIP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "MIP" 'MI
		SQLE="select lpad(MIP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "FAMILY_OUTPUT" 'FO
		SQLE="select lpad(FAMILY_OUTPUT_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "SCHEDULE" 'SC
		SQLE="select lpad(SCHEDULE_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "LABOUR" 'LA
		SQLE="select lpad(LABOUR_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "SCANNER" 'SC
		SQLE="select lpad(SCANNER_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "LOG" 'LO   
		SQLE="select lpad(LOG_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "GENERAL" 'LA
		SQLE="select lpad(GENERAL_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "TRANSACTION" 'LA
		SQLE="select  JOB_TRANSACTION_SEQ.nextval as SEQ from DUAL"
		
	case "STATION_GROUP_SEQ" 'LA
		SQLE="select  lpad(STATION_GROUP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "STATION_GROUP_STATION_SEQ" 'LA
		SQLE="select  lpad(STATION_GROUP_STATION_SEQ.nextval,8,'0') as SEQ from DUAL"
	
	case "STATION_GROUP_MAPPING_SEQ" 'LA
		SQLE="select  lpad(STATION_GROUP_MAPPING_SEQ.nextval,8,'0') as SEQ from DUAL"
		
	case "Part_New_SEQ" 'LA
		SQLE="select  lpad(PART_SEQUENCE.nextval,5,'0') as SEQ from DUAL"	
	end select

	rsE.open SQLE,conn,1,3
	SEQ=rsE("SEQ")
	rsE.close
NID_SEQ=SEQ
set rsE=nothing
end function
%>