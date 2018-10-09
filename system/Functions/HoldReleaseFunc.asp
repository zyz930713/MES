<!--#include virtual="/Functions/SendJMail.asp" -->
<%
function getHoldType(holdType)
    set rsGroup=server.CreateObject("adodb.recordset")    
    rsGroup.open "select nid,group_name,group_chinese_name from system_group where group_type='Hold' ",conn,1,3
    while not rsGroup.eof
        output = output&"<option value='"&rsGroup("nid")&"' "
		if holdType=rsGroup("nid") then
			 output = output &" selected "     
		end if
		output = output & ">"&rsGroup("group_name")&" ("&rsGroup("group_chinese_name")&")</option>"  
        rsGroup.movenext
	wend	
    rsGroup.close
    set rsGroup=nothing
    getHoldType=output
end function


 'add by Lennie Wei 2013-01-10
function sendHoldReleaseNotiInfo(jobNum,actionType,reason,holdId,partNum)
    on error resume next
    mailTo = ""
    mailContents = "<style>.table_th {color:#FFFFFF;background-color:#73A2EE;font-size: 16px;height:20px;width:120px}</style><meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />"
    mailContents = mailContents & "Hi all,<br><table><tr><td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td>These jobs have been"    
    if actionType = "Hold" then
        mailContents = mailContents & " held "
    else 
        mailContents = mailContents & " released "
    end if
    mailContents = mailContents & "by "& session("user") &"(" &session("code") &"). </td></tr>"
    mailContents = mailContents & "<tr><td></td><td>The reason is: " & reason &".</tr><tr><td></td><td>Job list:</td>"
    mailContents = mailContents & "</tr><tr><td></td><td><table border='1' cellpadding='0' cellspacing='0' ><th class='table_th'> Job Number </th><th class='table_th'> Part Number </th>"
    
    job = split(jobNum,";")
	part = split(partNum,";")
    for i=0 to ubound(job)
        mailContents = mailContents & "<tr><td>" & job(i) & "</td><td>" & part(i) & "</td></tr>"
    next
    
    mailContents = mailContents & "</table></td></tr></table><p>Best Regards!<p>"&application("SystemName") 
    
    'get hold info
    holdPerson = ""
    holdType = ""
    set rsHold=server.CreateObject("adodb.recordset")
    holdId = "'" & replace(holdId,",","','") & "'"    
    strSql = "select distinct transaction_person,hold_type_id from job_hold_release_history where id in (" & holdId & ")"
    rsHold.open strSql,conn,1,3
    while not rsHold.eof
        holdPerson = holdPerson &"," & rsHold("transaction_person")
        holdType = holdType & "," & rsHold("hold_type_id")
        rsHold.movenext
    wend
    rsHold.close
    set rsHold=nothing
	
    set rsMail=server.CreateObject("adodb.recordset")
    strSql = "select group_members as email from system_group where nid in ('" & replace(holdType,",","','") & "') "
    strSql = strSql & " union "
    strSql = strSql & " select email from users where user_code in ('" & replace(holdPerson,",","','") & "') "
    
    rsMail.open strSql,conn,1,3
    while not rsMail.eof
        mailTo = mailTo  & replace(rsMail(0),",",";") & ";"
        rsMail.movenext
    wend
    rsMail.close
    set rsMail=nothing

    if mailTo <> "" then	 
		mailTo=left(mailTo,len(mailTo)-1)	               
        SendJMail application("MailSender"),mailTo,actionType & " Job Notification",mailContents
    end if

end function

 %>