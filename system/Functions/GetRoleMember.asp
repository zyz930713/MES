<%
function getRoleMember(isshowname,showstyle,role,where,order,sequency,splitchar)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
set rsE=server.CreateObject("adodb.recordset")
SQLE="Select USER_CODE,USER_NAME from USERS "& where
rsE.open SQLE,conn,1,3
if not rsE.eof then
z=rsE.recordcount
	if sequency<>"" then
		aseq=split(sequency,",")
		for x=0 to ubound(aseq)
			while not rsE.eof
				if aseq(x)=rsE("USER_CODE") then
					select case showstyle
					case "OPTION"
					output=output&"<option value='"&rsE("USER_NAME")&"'"
						if rsE("USER_CODE")=station then
						output=output&"selected"
						end if
					output=output&">"&rsE("USER_NAME")&"</option>"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsE("USER_NAME")&splitchar
						else
						output=output&rsE("USER_NAME")
						end if
					case "CODE"
						if x<>ubound(aseq) then
						output=output&rsE("USER_CODE")&splitchar
						else
						output=output&rsE("USER_CODE")
						end if
					case "CHECK"
						output=output&"<input type=""checkbox"" name=""members"" value="""&rsE("USER_CODE")&""" checked>"&rsE("USER_NAME")
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
				output=output&"<option value='"&rsE("USER_NAME")&"'"
					if rsE("USER_CODE")=role then
					output=output&"selected"
					end if
				output=output&">"&rsE("USER_NAME")&"</option>"
			case "TEXT"
				if y<>z then
				output=output&rsE("USER_NAME")&splitchar
				else
				output=output&rsE("USER_NAME")
				end if
			case "CODE"
				if y<>z then
				output=output&rsE("USER_CODE")&splitchar
				else
				output=output&rsE("USER_CODE")
				end if
			case "CHECK"
				output=output&"<input type=""checkbox"" name=""members"" value="""&rsE("USER_CODE")&""" checked>"&rsE("USER_NAME")
			end select
		rsE.movenext
		y=y+1
		wend
	end if
end if
getRoleMember=output
rsE.close
set rsE=nothing
end function
%>