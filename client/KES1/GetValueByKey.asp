<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=application("SystemName")%></title>
</head>
<%
key = request("key")
keyValue = request("keyValue")
rtnValue=""

'get job number by tray id
if key = "TrayId" then 
	strSql = "select job_number,sheet_number from JOB_TRAY_MAPPING where tray_id= '"& keyValue &"' And is_disable = 0 "
	rs.open strSql,conn,1,3
	if rs.recordcount=1 then
		rtnValue = rs("job_number") &"-"& string(3-len(rs("sheet_number")),"0")&rs("sheet_number")
	elseif rs.recordcount > 1 then
		rtnValue = "Error-This tray id has linked more than one job number.\n该料盘与工单的对应关系不唯一。"
	else
		rtnValue = "Error-This tray id has not lined a job number.\n该料盘没有建立与工单的对应关系。"
	end if
	rs.close
'get defect code info	
elseif key = "DefectCode" then
	strSql = "select nid,defect_name||'('||defect_chinese_name||')' as defect_name,station_id,transaction_type from defectcode where routing_id = "
	strSql = strSql+"(select mother_routing_id from part where nid='"&session("PART_NUMBER_ID")&"') and defect_code='"& keyValue & "' "		
	rs.open strSql,conn,1,3
	  rtnValue="sdfsdf"
	stationIndex=0
	if rs.eof then
		rtnValue = "Error-This defect code is not in the routing of this job number.\n该工单的制程中不包含此不良代码。"
	else	
		while not rs.eof
			if instr(session("DEFECT_STATIONS"),rs("station_id"))>stationIndex then
				stationIndex = instr(session("DEFECT_STATIONS"),rs("station_id"))
				rtnValue = rs("nid")&"$"&rs("defect_name")&"$"&rs("station_id")&"$"&rs("transaction_type")
			end if
			rs.movenext
		wend	
		if stationIndex=0 then
			rtnValue = "Error-This defect code was not occured in the finished stations or current station. \n此不良代码不是已完成站点或当前站点产生的。"			
		end if
	end if
'check 2D Code	
elseif key = "2DCode" then
	strSql = "select code from job_2d_code where code in ('"& replace(keyValue,",","','") & "') "
	rs.open strSql,conn,1,3
'	if not rs.eof then
'		rtnValue = "Error-This 2D code("+keyValue+") is existed.\n该二维码("+keyValue+")已经存在。"
'	end if
	while not rs.eof
		rtnValue=rtnValue&rs("code")&","
		rs.movenext
	wend
	rs.close	
'get sub job by computer
elseif key = "computer" then
	strSql = "select job_number,sheet_number from job a,computer_printer_mapping b where instr(b.line_name,a.line_name)>0 "
	strSql = strSql + " and a.status=0 and b.computer_name='"& trim(keyValue) & "' and a.start_time >trunc(sysdate) "
	strSql = strSql + " and exists(select 1 from station where (instr(a.finished_stations_id,nid)=0 or a.finished_stations_id is null) and mother_station_id='"&Request.Cookies("current_station_id")&"') "
	strSql = strSql + " order by a.start_time"
  
	rs.open strSql,conn,1,3
	if not rs.eof then
		rtnValue = rs("job_number") &"-"& string(3-len(rs("sheet_number")),"0")&rs("sheet_number")
	end if
	rs.close	
elseif key="Defect2DCode" then
	strSql = "select defect_code_id,lm_user,lm_time from job_2d_code where code = '"& trim(keyValue) & "' and job_number='"&session("JOB_NUMBER")&"' and sheet_number='"&session("SHEET_NUMBER")&"' "
	
	rs.open strSql,conn,1,3
	if not rs.eof then
		if rs("defect_code_id") <> "" then
			rtnValue = "Error-This 2D code("+keyValue+") was bound defect.\n该二维码("+keyValue+")已经绑定了不良。"
		else
			'rs("defect_code_id")=request("defect_id")
			'rs("lm_user")=session("code")
			'rs("lm_time")=now()
			'rs.update
		end if		
	else		
		rtnValue = "Error-This 2D code("+keyValue+") is not belong the job number.\n该二维码("+keyValue+")不输入此工单。"
	end if
	rs.close
elseif key="JobNumber" then

  strSql="select * from REPORT_HV_DATA where job_number='"&keyValue&"'"
    rs.open strSql,conn,1,3
	if not rs.eof then
	rtnValue="OK"
	
	end if
	
	rs.close
end if
%>
<script type="text/javascript">
	window.returnValue="<%=rtnValue%>";
	window.close();
</script>
<body>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
