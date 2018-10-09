<%
'Function GetUserReportTo(UserCode,ReturnValue)
'	set oRs=server.CreateObject("adodb.recordset")
'	set objRst=server.CreateObject("adodb.recordset")
'	SQL="select * from SYSTEM_USERS where USER_CODE='" & UserCode & "'"
'	
'	oRs.open SQL,conn_web,1,3
'	
'	if not oRs.eof then
'		
'		SQL="select * from SYSTEM_USERS where USER_CODE='" & trim(oRs("REPORT_TO")&"") & "'"
'		objRst.open SQL,conn_web,1,3
'		
'		if not objRst.eof then
'			'if trim(objRst("agent")&"")="" then
'				ReportToCode=trim(oRs("REPORT_TO")&"")
'			'else
'			'	ReportToCode=trim(objRst("agent")&"")
'			'end if
'		end if
'		objRst.close
'		select case ReturnValue
'			case "code"
'				GetUserReportTo=ReportToCode
'			case "ename"
'				set oRst=server.CreateObject("adodb.recordset")
'				SQL="select * from SYSTEM_USERS where USER_CODE='" & ReportToCode & "'"
'				oRst.open SQL,conn_web,1,3
'				
'				if not oRst.eof then
'					GetUserReportTo=trim(oRst("ENGLISH_NAME")&"")
'				else
'					GetUserReportTo=""
'				end if
'				oRst.close
'				set oRst=nothing
'			case "cname"
'				set oRst=server.CreateObject("adodb.recordset")
'				SQL="select * from SYSTEM_USERS where USER_CODE='" & ReportToCode & "'"
'				oRst.open SQL,conn_web,1,3
'				
'				if not oRst.eof then
'					GetUserReportTo=trim(oRst("CHINESE_NAME")&"")
'				else
'					GetUserReportTo=""
'				end if
'				oRst.close
'		end select
'	else
'	
'		GetUserReportTo=""
'		
'	end if
'	oRs.close
'	set oRs=nothing
'	set objRst=nothing
'End Function

Function GetUserReportTo(UserCode,ReturnValue)
	set oRs=server.CreateObject("adodb.recordset")
	set objRst=server.CreateObject("adodb.recordset")
	SQL="select * from PERSONNEL01 where CODE='" & UserCode & "'"
	
	oRs.open SQL,conn_web,1,3
	
	if not oRs.eof then
		
		SQL="select * from PERSONNEL01 where CODE='" & trim(oRs("ReportTo")&"") & "'"
		objRst.open SQL,conn_web,1,3
		
		if not objRst.eof then
			'if trim(objRst("agent")&"")="" then
				ReportToCode=trim(oRs("ReportTo")&"")
			'else
			'	ReportToCode=trim(objRst("agent")&"")
			'end if
		end if
		objRst.close
		select case ReturnValue
			case "code"
				GetUserReportTo=ReportToCode
			case "ename"
				set oRst=server.CreateObject("adodb.recordset")
				SQL="select * from PERSONNEL01 where CODE='" & ReportToCode & "'"
				oRst.open SQL,conn_web,1,3
				
				if not oRst.eof then
					GetUserReportTo=trim(oRst("EnglishName")&"")
				else
					GetUserReportTo=""
				end if
				oRst.close
				set oRst=nothing
			case "cname"
				set oRst=server.CreateObject("adodb.recordset")
				SQL="select * from PERSONNEL01 where CODE='" & ReportToCode & "'"
				oRst.open SQL,conn_web,1,3
				
				if not oRst.eof then
					GetUserReportTo=trim(oRst("NAME")&"")
				else
					GetUserReportTo=""
				end if
				oRst.close
		end select
	else
	
		GetUserReportTo=""
		
	end if
	oRs.close
	set oRs=nothing
	set objRst=nothing
End Function

Function GetCurrentManager(UserCode,ReturnValue)
	set oRs=server.CreateObject("adodb.recordset")
	set objRst=server.CreateObject("adodb.recordset")
	SQL="select * from PERSONNEL01 where CODE='" & UserCode & "'"
	
	oRs.open SQL,conn_web,1,3
	
	if not oRs.eof then
		
		SQL="select * from PERSONNELWEB where CODE='" & trim(oRs("ReportTo")&"") & "'"
		objRst.open SQL,conn_web,1,3
		
		if not objRst.eof then
			if trim(objRst("agent")&"")="" then
				ReportToCode=trim(oRs("ReportTo")&"")
			else
				ReportToCode=trim(objRst("agent")&"")
			end if
		end if
		objRst.close
		select case ReturnValue
			case "code"
				GetCurrentManager=ReportToCode
			case "ename"
				set oRst=server.CreateObject("adodb.recordset")
				SQL="select * from PERSONNEL01 where CODE='" & ReportToCode & "'"
				oRst.open SQL,conn_web,1,3
				
				if not oRst.eof then
					GetCurrentManager=trim(oRst("EnglishName")&"")
				else
					GetCurrentManager=""
				end if
				oRst.close
				set oRst=nothing
			case "cname"
				set oRst=server.CreateObject("adodb.recordset")
				SQL="select * from PERSONNEL01 where CODE='" & ReportToCode & "'"
				oRst.open SQL,conn_web,1,3
				
				if not oRst.eof then
					GetCurrentManager=trim(oRst("NAME")&"")
				else
					GetCurrentManager=""
				end if
				oRst.close
		end select
	else
	
		GetCurrentManager=""
		
	end if
	oRs.close
	set oRs=nothing
	set objRst=nothing
End Function
%>