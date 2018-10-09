<%
function  GetTaskParamHTML(paramindex,paramtype,paramname,cparamname,paramvalue,thisoutput,thiscoutput)
	select case paramtype
	case "text"
	thisoutput="<input name='param"&paramindex&"' type='text' value='"&paramvalue&"'/>&nbsp;"&paramname
	thiscoutput="<input name='param"&paramindex&"' type='text' value='"&paramvalue&"'/>&nbsp;"&cparamname
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

	case "checkbox"
		if paramname="[FACTORY]" then
			FactoryRight ""
			set rsTPP=server.CreateObject("adodb.recordset")
			SQLTPP="select * from FACTORY"&factorywhereinside
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
	case "period"
		if paramvalue<>"" then
		a_value=split(paramvalue,"|")
		turn_value=a_value(1)
		week_value=a_value(2)
		hour_value=hour(3)
		minute_value=minute(4)
		else
		turn_value="0"
		week_value=weekday(date())
		hour_value=hour(now())
		minute_value=minute(now())
		end if
		if paramname="[FROM]" then
			thisoutput="From"
			thiscoutput="从"
		elseif paramname="[TO]" then
			thisoutput="To"
			thiscoutput="到"
		end if
		thisoutput=thisoutput&"<input name='param_type"&paramindex&"' type='hidden' id='param_type"&paramindex&"' value='"&paramtype&"'>&nbsp;<select name='param"&paramindex&"'><option>-- Select Turn --</option><option value='-1'"
		thiscoutput=thiscoutput&"<input name='param_type"&paramindex&"' type='hidden' id='param_type"&paramindex&"' value='"&paramtype&"'>&nbsp;<select name='param"&paramindex&"'><option>-- 选择周期 --</option><option value='-1'"
		if turn_value="-1" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Last</option><option value='0'"
		thiscoutput=thiscoutput&">上周</option><option value='0'"
		if turn_value="0" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Current</option><option value='1'"
		thiscoutput=thiscoutput&">本周</option><option value='1'"
		if turn_value="1" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Next</option></select><select name='param"&paramindex&"'><option>-- Select Weekday --</option><option value='2'"
		thiscoutput=thiscoutput&">下周</option></select><select name='param"&paramindex&"'><option>-- 选择星期 --</option><option value='2'"
		if week_value="2" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Monday</option><option value='3'"
		thiscoutput=thiscoutput&">星期一</option><option value='3'"
		if week_value="3" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Tuesday</option><option value='4'"
		thiscoutput=thiscoutput&">星期二</option><option value='4'"
		if week_value="4" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Wednesday</option><option value='5'"
		thiscoutput=thiscoutput&">星期三</option><option value='5'"
		if week_value="5" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Tuesday</option><option value='6'"
		thiscoutput=thiscoutput&">星期四</option><option value='6'"
		if week_value="6" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Friday</option><option value='7'"
		thiscoutput=thiscoutput&">星期五</option><option value='7'"
		if week_value="7" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Saturday</option><option value='1'"
		thiscoutput=thiscoutput&">星期六</option><option value='1'"
		if week_value="1" then
		thisoutput=thisoutput&" selected"
		thiscoutput=thiscoutput&" selected"
		end if
		thisoutput=thisoutput&">Sunday</option></select><input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&hour_value&"' size='2' onChange='hourcheck(this)'>:<input name='param"&paramindex&"' type='text' id='fromminute' value='"&minute_value&"' size='2' onChange='minutecheck(this)'><"
		thiscoutput=thiscoutput&">星期日</option></select><input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&hour_value&"' size='2' onChange='hourcheck(this)'>:<input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&minute_value&"' size='2' onChange='minutecheck(this)'><"
	case "time"
		if paramvalue<>"" then
		a_value=split(paramvalue,":")
		hour_value=trim(a_value(0))
		minute_value=trim(a_value(1))
		else
		hour_value=hour(now())
		minute_value=minute(now())
		end if
		if paramname="[FROM]" then
			thisoutput="From"
			thiscoutput="从"
		elseif paramname="[TO]" then
			thisoutput="To"
			thiscoutput="到"
		end if
		thisoutput=thisoutput&"<input name='param_type"&paramindex&"' type='hidden' id='param_type"&paramindex&"' value='"&paramtype&"'><input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&hour_value&"' size='2' onChange='hourcheck(this)'>:<input name='param"&paramindex&"' type='text' id='fromminute' value='"&minute_value&"' size='2' onChange='minutecheck(this)'><"
		thiscoutput=thiscoutput&"<input name='param_type"&paramindex&"' type='hidden' id='param_type"&paramindex&"' value='"&paramtype&"'><input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&hour_value&"' size='2' onChange='hourcheck(this)'>:<input name='param"&paramindex&"' type='text' id='param"&paramindex&"' value='"&minute_value&"' size='2' onChange='minutecheck(this)'><"	
	end select
end function
%>