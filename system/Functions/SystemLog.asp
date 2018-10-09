<%
function SystemLog(module,action)
set rsSys=server.CreateObject("adodb.recordset")
SQLSys="select * from SYSTEM_LOG"
rsSys.open SQLSys,conn,3,3
rsSys.addnew
rsSys("NID")="LO"&NID_SEQ("LOG")
rsSys("OCCURRED_TIME")=now()
rsSys("USER_CODE")=session("code")
rsSys("USER_NAME")=session("user")
rsSys("MODULE")=module
rsSys("ACTION")=action
rsSys.update
rsSys.close
set rsSys=nothing
end function
%>