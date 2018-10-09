<%
function  GetFormParamHTML(paramindex,paramtype,paramname,cparamname,paramscripts,paramvalue,paramshowbutton,cparamshowbutton,parambuttonscripts,paramtitle,cparamtitle,thisoutput,thiscoutput,jobnumber)
	outputhead="<input name='paramname"&paramindex&"' type='hidden' id='paramname"&paramindex&"' value='"&paramname&"'><input name='paramindex"&paramindex&"' type='hidden' id='paramindex"&paramindex&"' value='"&paramindex&"'>"
	select case paramtype
	case "text"
	if paramname="Job Number" and jobnumber<>"" then
	paramvalue=jobnumber
	end if
	thisoutput=paramname&"&nbsp;<input name='param"&paramindex&"' type='text' value='"&paramvalue&"' onChange='"&paramscripts&"'/>"
	thiscoutput=cparamname&"&nbsp;<input name='param"&paramindex&"' type='text' value='"&paramvalue&"' onChange='"&paramscripts&"'/>"
	if paramshowbutton<>"" then
	thisoutput=thisoutput&"<input name='Show' type='button' value='"&paramshowbutton&"' onClick='"&parambuttonscripts&"'/>"
	thiscoutput=thiscoutput&"<input name='Show' type='button' value='"&cparamshowbutton&"' onClick='"&parambuttonscripts&"'/>"
	end if
	if paramtitle<>"" then
	thisoutput=thisoutput&"&nbsp;"&paramtitle
	thiscoutput=thiscoutput&"&nbsp;"&cparamtitle
	end if
	case "radio"
		radiotext=split(paramname,";")
		cradiotext=split(cparamname,";")
		for i=0 to ubound(radiotext)
			if cstr(i+1)=paramvalue then
			thisoutput=thisoutput&"<input name='param"&paramindex&"' type='radio' value='"&i+1&"' checked/>"&radiotext(i)
			else
			thisoutput=thisoutput&"<input name='param"&paramindex&"' type='radio' value='"&i+1&"'/>"&radiotext(i)
			end if
		next
		for i=0 to ubound(cradiotext)
			if cstr(i+1)=paramvalue then
			thiscoutput=thiscoutput&"<input name='param"&paramindex&"' type='radio' value='"&i+1&"' checked/>"&cradiotext(i)
			else
			thiscoutput=thiscoutput&"<input name='param"&paramindex&"' type='radio' value='"&i+1&"'/>"&cradiotext(i)
			end if
		next
	case "option"
		if instr(lcase(paramname),"line")>0 then
		thisoutput="<select name='param"&paramindex&"' id='param"&paramindex&"'><option value=''>--Select--</option></select>"
		end if
		if instr(lcase(paramname),"line")>0 then
		thiscoutput="<select name='param"&paramindex&"' id='param"&paramindex&"'><option value=''>--Ñ¡Ôñ--</option></select>"
		end if
	case "checkbox"
		if paramname="[FACTORY]" then
			FactoryRight '
			set rsTPP=server.CreateObject("adodb.recordset")
			SQLTPP="select * from FACTORY"&factorywhereoutside
			rsTPP.open SQLTPP,conn,1,3
			if not rsTPP.eof then
				while not rsTPP.eof
					if instr(paramvalue,rsTPP("NID"))>0 then
					thisoutput=thisoutput&"<input name='param"&paramindex&"' type='checkbox' id='param"&paramindex&"' value='"&rsTPP("NID")&"' checked>"&rsTPP("FACTORY_NAME")
					else
					thisoutput=thisoutput&"<input name='param"&paramindex&"' type='checkbox' id='param"&paramindex&"' value='"&rsTPP("NID")&"'>"&rsTPP("FACTORY_NAME")
					end if				
				rsTPP.movenext
				wend
			end if	
			rsTPP.close
			set rsTPP=nothing
			thiscoutput=thisoutput
		end if
	case "calendar"
		thisoutput=paramname&"&nbsp;<input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&paramvalue&"' size='10'>"
		thiscoutput=cparamname&"&nbsp;<input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&paramvalue&"' size='10'>"
	end select
	thisoutput=outputhead&thisoutput
	thiscoutput=outputhead&thiscoutput
end function
%>