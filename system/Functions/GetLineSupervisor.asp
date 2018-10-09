<%
function getLineSupervisor(showstyle,line,where,order,splitchar)
set rsL=server.CreateObject("adodb.recordset")
SQLP="Select 1,L.NID,L.LINE_NAME from LINE L inner join FACTORY F on L.FACTORY_ID=F.NID "&where&order
rsL.open SQLP,conn,1,3
if not rsL.eof then
	while not rsL.eof
		select case showstyle
		case "OPTION"
		output=output&"<option value='"&rsL("NID")&"'"
			if rsL("NID")=line then
			output=output&"selected"
			end if
		output=output&">"&rsL("LINE_NAME")&"</option>"
		case "OPTION_LINE_NAME"
		output=output&"<option value='"&rsL("LINE_NAME")&"'"
			if rsL("LINE_NAME")=line then
			output=output&"selected"
			end if
		output=output&">"&rsL("LINE_NAME")&"</option>"
		case "RADIO"
		output=output&"<input name='factory' type='radio' class='t-c-greenCopy' value='"&rsL("NID")&"'>"&rsL("LINE_NAME")&"&nbsp;"
		case "TEXT"
			output=output&rsL("LINE_NAME")&splitchar
		end select
	rsL.movenext
	wend
end if
getLine=output
rsL.close
set rsL=nothing
end function
%>