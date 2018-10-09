<%
function GetWelcomeSettingsByDay(thisday,path,query)
SQL="select * from DAILY_WELCOME where to_char(WELCOME_DAY,'yyyy-mm-dd')='"&formatdate(thisday,application("F_shortdateformat"))&"'"
rs.open SQL,conn,1,3
GetWelcomeSettingsByDay="<table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
if not rs.eof then
	while not rs.eof
		if rs("WELCOME_TYPE")="2" and rs("FILE_NAME")<>"" then
		GetWelcomeSettingsByDay=GetWelcomeSettingsByDay&"<tr><td><input name=""type"&replace(thisday,"-","_")&""" type=""radio"" value=""1"" onClick=""typecheck('"&replace(thisday,"-","_")&"','1','"&path&"','"&query&"')""/></td><td><div align=""left"">Barcode</div></td></tr><tr><td><input name=""type"&replace(thisday,"-","_")&""" type=""radio"" value=""2"" onClick=""typecheck('"&replace(thisday,"-","_")&"','2','"&path&"','"&query&"')"" checked/></td><td><div align=""left"" class=""red"">"&rs("FILE_NAME")
		else
		GetWelcomeSettingsByDay=GetWelcomeSettingsByDay&"<tr><td><input name=""type"&replace(thisday,"-","_")&""" type=""radio"" value=""1"" onClick=""typecheck('"&thisday&"','1','"&path&"','"&query&"')"" checked/></td><td><div align=""left"">Barcode</div></td></tr><tr><td><input name=""type"&replace(thisday,"-","_")&""" type=""radio"" value=""2"" onClick=""typecheck('"&thisday&"','2','"&path&"','"&query&"')"" /></td><td>Powerpoint"
		end if
		GetWelcomeSettingsByDay=GetWelcomeSettingsByDay&"</div>&nbsp;<span style=""cursor:hand"" onClick=""location.href='/Admin/LCD/Welcome.asp?thisday="&thisday&"&path="&path&"&query="&query&"'"" title=""Upload Powerpoint File"">&nbsp;<img src=""/Images/IconUpload.gif"" align=""absmiddle"" /></span></td></tr>"
	rs.movenext
	wend
else
	GetWelcomeSettingsByDay=GetWelcomeSettingsByDay&"<tr><td><input name=""type"&replace(thisday,"-","_")&""" type=""radio"" value=""1"" onClick=""typecheck('"&replace(thisday,"-","_")&"','1','"&path&"','"&query&"')"" checked /></td><td><div align=""left"">Barcode</div></td></tr><tr><td><input name=""type"&replace(thisday,"-","_")&""" type=""radio"" value=""2"" onClick=""typecheck('"&replace(thisday,"-","_")&"','2','"&path&"','"&query&"')"" /></td><td>Powerpoint<span style=""cursor:hand"" onClick=""location.href='/Admin/LCD/Welcome.asp?thisday="&thisday&"&path="&path&"&query="&query&"'"" title=""Upload Powerpoint File"">&nbsp;<img src=""/Images/IconUpload.gif"" align=""absmiddle"" /></span></td></tr>"
end if
rs.close
GetWelcomeSettingsByDay=GetWelcomeSettingsByDay&"</table>"
end function
%>