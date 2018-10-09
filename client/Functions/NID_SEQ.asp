<%
function NID_SEQ(table)
set rsE=server.CreateObject("adodb.recordset")
	select case table
	case "PROD_STORE" 'WH
		SQLE="select lpad(STORE_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "PROD_SCRAP" 'CH
		SQLE="select lpad(SCRAP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "STOCK" 'WP for KE (Store List)
		SQLE="select lpad(STOCK_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "EMC_STOCK" 'WE for EMC (Store List)
		SQLE="select lpad(EMC_STOCK_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "VAM_STOCK" 'WV for VAM (Store List)
		SQLE="select lpad(VAM_STOCK_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "KE_SCRAP" 'SP for KE (Scrap List)
		SQLE="select lpad(KE_SCRAP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "EMC_SCRAP" 'SE for EMC (Scrap List)
		SQLE="select lpad(EMC_SCRAP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "VAM_SCRAP" 'SV for VAM (Scrap List)
		SQLE="select lpad(VAM_SCRAP_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "CMMSJOB" 'CJ
		SQLE="select lpad(KESCMMS.JOB_SEQ.nextval,4,'0') as SEQ from DUAL"
	case "Retest" 'RT Retest_seq  
		SQLE="select lpad(Retest_seq.nextval,8,'0') as SEQ from DUAL"
	case "Retest_Detail" 'RT Retest_seq  RETEST_DETAIL_SEQ
		SQLE="select lpad(RETEST_DETAIL_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "Retest_Defect" 'RT Retest_seq  RETEST_DETAIL_SEQ
		SQLE="select lpad(RETEST_DEFECT_SEQ.nextval,8,'0') as SEQ from DUAL"
	
	
	case "OQC_SEQ" 'RT Retest_seq  RETEST_DETAIL_SEQ
		SQLE="select lpad(OQC_SEQ.nextval,8,'0') as SEQ from DUAL"
	case "OQC_DETAIL_SEQ" 'RT Retest_seq  RETEST_DETAIL_SEQ
		SQLE="select lpad(OQC_DETAIL_SEQ.nextval,8,'0') as SEQ from DUAL"
		
	end select
	rsE.open SQLE,conn,1,3
	SEQ=rsE("SEQ")
	rsE.close
NID_SEQ=SEQ
set rsE=nothing
end function
%>