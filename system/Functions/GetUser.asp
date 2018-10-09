<%
function getUser(isshowname,showstyle,user,where,order,sequency,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsE=server.CreateObject("adodb.recordset")
SQLE="Select NID,USER_NAME,USER_CODE from USERS "& where&order
rsE.open SQLE,conn,1,3
if not rsE.eof then
z=rsE.recordcount
	if sequency<>"" then
		aseq=split(sequency,",")
		for x=0 to ubound(aseq)
			while not rsE.eof
				if aseq(x)=rsE("NID") then
					select case showstyle
					case "OPTION"
					output=output&"<option value='"&rsE("NID")&"'"
						if rsE("NID")=user then
						output=output&"selected"
						end if
					output=output&">"&rsE("USER_NAME")&"("&rsE("USER_CODE")&")</option>"
					case "USER_GROUP_OPTION"
					output=output&"<option value='"&rsE("USER_CODE")&"'"
						if rsE("USER_CODE")=user then
						output=output&"selected"
						end if
					output=output&">"&rsE("USER_NAME")&"("&rsE("USER_CODE")&")</option>"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsE("USER_NAME")&splitchar
						else
						output=output&rsE("USER_NAME")
						end if
					end select
				end if
			rsE.movenext
			wend
		rsE.movefirst
		next
	else
		y=1
		while not rsE.eof
			select case showstyle
			case "OPTION"
				output=output&"<option value='"&rsE("NID")&"'"
					if rsE("NID")=user then
					output=output&"selected"
					end if
				output=output&">"&rsE("USER_NAME")&"("&rsE("USER_CODE")&")</option>"
			case "USER_GROUP_OPTION"
				output=output&"<option value='"&rsE("USER_CODE")&"'"
					if rsE("USER_CODE")=user then
					output=output&"selected"
					end if
				output=output&">"&rsE("USER_NAME")&"("&rsE("USER_CODE")&")</option>"
			case "TEXT"
				if y<>z then
				output=output&rsE("USER_NAME")&splitchar
				else
				output=output&rsE("USER_NAME")
				end if
			end select
		rsE.movenext
		y=y+1
		wend
	end if
end if
getUser=output
rsE.close
set rsE=nothing
end function
%>