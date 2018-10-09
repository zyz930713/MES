<%
Function GetUserInfo(UserCode,ReturnItem)
	set objRst=server.CreateObject("adodb.recordset")
	
	if ReturnItem="USER_CODE" then
		SQL="SELECT * FROM PERSONNEL01 WHERE EnglishName='" & UserCode & "'"
	else
		if ReturnItem="EMAIL" then
			SQL="SELECT * FROM PERSONNELWEB WHERE CODE='" & UserCode & "'"
		else		
			SQL="SELECT * FROM PERSONNEL01 WHERE (CODE='" & UserCode & "' or EnglishName='" & UserCode & "')"
		END IF
	end if
	objRst.open SQL,conn_web,1,3
	if not objRst.eof then
		GetUserInfo=objRst(ReturnItem)
	else
		GetUserInfo=""
	end if
	objRst.close
	set objRst=nothing
End Function
%>