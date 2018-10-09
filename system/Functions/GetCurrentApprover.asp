<%
Function GetCurrentApprover(formcode)
	set oRs=server.CreateObject("adodb.recordset")
	SQL="select * from EFORM_FORM_QUEUE where FormCode='" & formcode & "'"
	oRs.open SQL,conn,1,3
	if not oRs.EOF then
		strCurrentApprover=oRs("CURRENT_APPROVER")
	end if
	oRs.close
	
'	SQL="select AGENT from SYSTEM_USERS where USER_CODE='" & strCurrentApprover & "'"
'	oRs.open SQL,conn,1,3
'	if not oRs.eof then
'		if trim(oRs("AGENT")&"")<>"" then
'			GetCurrentApprover=trim(oRs("AGENT")&"")
'		else
'			GetCurrentApprover=strCurrentApprover
'		end if
'
'	end if
'	oRs.close	
	GetCurrentApprover=strCurrentApprover	
	 
End Function

Function GetApproverList(formcode)

	set oRs=server.CreateObject("adodb.recordset")
	SQL="select * from EFORM_FORM_QUEUE where FormCode='" & formcode & "'"
	oRs.open SQL,conn,1,3
	if not oRs.EOF then
		strApproverList=oRs("APPROVER_CODE")
	end if
	oRs.close
	
	GetApproverList=strApproverList
	 
End Function

Function GetApproverListHTML(formcode)

	strApproverListHTML=""
	strSelfHTML=""
	set oRs=server.CreateObject("adodb.recordset")
	SQL="select * from EFORM_FORM_QUEUE where FormCode='" & formcode & "'"
	oRs.open SQL,conn,1,3
	if not oRs.EOF then
		INPUT_CODE=trim(oRs("INPUT_CODE")&"")
		arrCodes=split(oRs("APPROVER_CODE"),",")
		if trim(oRs("APPROVE_DATE")&"")<>"" then
			arrDates=split(oRs("APPROVE_DATE"),",")
		end if
		for y=0 to ubound(arrCodes)
			if trim(arrCodes(y)&"")<>"" then
				
				if session("language")=0 then
					if instr(oRs("PASS_CODE"),trim(arrCodes(y)&""))>0 then
						strApproverListHTML=strApproverListHTML&"<font color=""#006600"">" & GetUserInfo(trim(arrCodes(y)&""),"ENGLISH_NAME") & "(" & arrDates(y) & ") ¡ú </font>"
					else
						if trim(arrCodes(y)&"")=trim(oRs("CURRENT_APPROVER")&"") then
							strApproverListHTML=strApproverListHTML&"<font color=""#FF0000"">" & GetUserInfo(trim(arrCodes(y)&""),"ENGLISH_NAME") & " ¡ú </font>"
						else
							strApproverListHTML=strApproverListHTML&"<font color=""#0000FF"">" & GetUserInfo(trim(arrCodes(y)&""),"ENGLISH_NAME")& " ¡ú </font>"
						end if
					end if				
				else
				
					if instr(oRs("PASS_CODE"),trim(arrCodes(y)&""))>0 then
						strApproverListHTML=strApproverListHTML&"<font color=""#006600"">" & GetUserInfo(trim(arrCodes(y)&""),"CHINESE_NAME") & "(" & arrDates(y) & ") ¡ú </font>"
					else
						if trim(arrCodes(y)&"")=trim(oRs("CURRENT_APPROVER")&"") then
							strApproverListHTML=strApproverListHTML&"<font color=""#FF0000"">" & GetUserInfo(trim(arrCodes(y)&""),"CHINESE_NAME") & " ¡ú </font>"
						else
							strApproverListHTML=strApproverListHTML&"<font color=""#0000FF"">" & GetUserInfo(trim(arrCodes(y)&""),"CHINESE_NAME")& " ¡ú </font>"
						end if
					end if
				end if
			end if
		next 	
		
	end if
	oRs.close
	if trim(strApproverListHTML&"")="" then 
		strApproverListHTML="&nbsp;"
	else
		strApproverListHTML=left(strApproverListHTML,(len(strApproverListHTML)-10))&"</font>"
		
	end if
	
	if INPUT_CODE<>"" then
		if GetReportToAgent(INPUT_CODE)=INPUT_CODE then
			if session("language")=0 then
				strSelfHTML="<font color=""#006600"">" & GetUserInfo(INPUT_CODE,"ENGLISH_NAME") & "(Agent for " & GetUserInfo(GetInputorReportCode(INPUT_CODE),"ENGLISH_NAME") & ") ¡ú </font>"
			else
				strSelfHTML="<font color=""#006600"">" & GetUserInfo(INPUT_CODE,"CHINESE_NAME") & "(´úÀí" & GetUserInfo(GetInputorReportCode(INPUT_CODE),"CHINESE_NAME") & ") ¡ú </font>"			
			end if
		end if
	end if
	GetApproverListHTML=strSelfHTML & strApproverListHTML
	 
End Function

Function GetReportToAgent(user_code)
	GetReportToAgent=""
	set objRs=server.CreateObject("adodb.recordset")
	SQL="select ReportTo from KESKQ30.dbo.Personnel01 where code='" & cstr(trim(user_code&"")) & "'"
	objRs.open SQL,conn,1,3
	if not objRs.eof then
		strReportTo=objRs("ReportTo")
	else
		strReportTo=""
	end if
	objRs.close
	if strReportTo<>"" then
		SQL="select agent from KESKQ30.dbo.PersonnelWeb where code='" & trim(strReportTo&"") & "'"
		objRs.open SQL,conn,1,3
		if not objRs.eof then
			GetReportToAgent=trim(objRs("agent")&"")
		end if
		objRs.close
	end if
	set objRs=nothing
End Function

Function GetInputorReportCode(user_code)
	GetInputorReportCode=""
	set objRs=server.CreateObject("adodb.recordset")
	SQL="select ReportTo from KESKQ30.dbo.Personnel01 where code='" & cstr(trim(user_code&"")) & "'"
	objRs.open SQL,conn,1,3
	if not objRs.eof then
		strReportTo=objRs("ReportTo")
	else
		strReportTo=""
	end if
	objRs.close

	set objRs=nothing
	GetInputorReportCode=trim(strReportTo&"")
End Function
%>