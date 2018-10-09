<%
const B_NUMBER=1
const B_STRING=2
const B_DATE=3
function getvalue(this_value,value_type)
	select case value_type
	case 1
		if this_value<>"" and isnull(this_value)=false then
		getvalue=csng(this_value)
		else
		getvalue=0
		end if
	end select
end function

function formatdate(vdate,vformat)
if vdate<>"" then
	vyear=year(vdate)
	vmonth=month(vdate)
	vday=day(vdate)
	
	vhour=hour(vdate)
	vminute=minute(vdate)
	vsecond=second(vdate)
	
	select case lcase(vformat)
	case "dd.mm.yyyy"
	formatdate=repeatstring(vday,"0",2)&"."&repeatstring(vmonth,"0",2)&"."& vyear
	case "yyyymmdd"
	formatdate=vyear&"-"&vmonth&"-"&vday
	case "@yyyymmdd@"
	formatdate=vyear&"-"&repeatstring(vmonth,"0",2)&"-"&repeatstring(vday,"0",2)
	case "ddmmyyyy"
	formatdate=vday&"/"&vmonth&"/"&vyear
	case "mmddyyyy"
	formatdate=vmonth&"/"&vday&"/"&vyear
	case "yyyymmdd hhnnss"
	formatdate=vyear&"-"&vmonth&"-"&vday&" "&vhour&":"&vminute&":"&vsecond
	case "ddmmyyyy hhnnss"
	formatdate=vday&"/"&vmonth&"/"&vyear&" "&vhour&":"&vminute&":"&vsecond
	case "mmddyyyy hhnnss"
	formatdate=vmonth&"/"&vday&"/"&vyear&" "&vhour&":"&vminute&":"&vsecond
	case "hhnnss"
	formatdate=vhour&":"&vminute&":"&vsecond
	case "hhnn"
	formatdate=vhour&":"&vminute
	case "hh"
	formatdate=vhour
	case else
	formatdate=vdate
	end select
end if
end function

function getmin(a,b)
	if a<b then
	getmin=a
	else
	getmin=b
	end if
end function

function repeatstring(str,ex,num)
	repeatstring=string(num-len(str),ex)&str
end function

function getStationTransaction(transaction,station_name,station_chinese_name)
	select case transaction
	case "2"
	getStationTransaction="["&station_name&"("&station_chinese_name&")]"
	case else
	getStationTransaction=station_name&"("&station_chinese_name&")"
	end select
end function

function formatlongstring(str,splitchar,num)
	mystr=""
	if  str<>"" then
		fj=1
		str=replace(str,"&nbsp;"," ")
		'str=replace(str,"<BR>","")
		for fi=1 to len(str)
			if fj>=num and mid(str,fi,1)=" "then
			fj=1
			mystr=mystr&mid(str,fi,1)&splitchar
			else
			fj=fj+1
			mystr=mystr&mid(str,fi,1)
			end if
		next
	end if
	formatlongstring=mystr
end function

function formatlongstringbreak(str,splitchar,num)
	mystr=""
	if  str<>"" then
		fj=1
		str=replace(str,"&nbsp;"," ")
		'str=replace(str,"<BR>","")
		for fi=1 to len(str)
			if fj>=num then
			fj=1
			mystr=mystr&mid(str,fi,1)&splitchar
			else
			fj=fj+1
			mystr=mystr&mid(str,fi,1)
			end if
		next
	end if
	formatlongstringbreak=mystr
end function

function weeknum(date1)
date0 = Year(date1) & "-1-1"    'firstday of year
weeknum = DateDiff("ww", date0, date1) + 1
end function

function excludesamestring(instring,splitstring)
	if instring<>"" then 
	stringarray=split(instring,splitstring)
		for fi=0 to ubound(stringarray)
			if instr(excludesamestring,stringarray(fi))<=0 then
			excludesamestring=excludesamestring&stringarray(fi)&","
			end if
		next
	excludesamestring=left(excludesamestring,len(excludesamestring)-1)
	else
	excludesamestring=""
	end if
end function

function highlightsamestring(instring,splitstring)
	if instring<>"" then 
	stringarray=split(instring,splitstring)
		for fi=0 to ubound(stringarray)
			if instr(highlightsamestring,stringarray(fi))>0 then
			highlightsamestring=highlightsamestring&"<span class='red'>&lt;"&stringarray(fi)&"&gt;</span>,"
			else
			highlightsamestring=highlightsamestring&stringarray(fi)&","
			end if
		next
	highlightsamestring=left(highlightsamestring,len(highlightsamestring)-1)
	else
	highlightsamestring=""
	end if
end function

function showsamestring(instring,splitstring)
	if instring<>"" then 
	stringarray=split(instring,splitstring)
		for fi=0 to ubound(stringarray)
			if instr(showsamestring,stringarray(fi))>0 then
			showsamestring=showsamestring&stringarray(fi)&","
			end if
		next
		if showsamestring<>"" then
		showsamestring=left(showsamestring,len(showsamestring)-1)
		end if
	else
	showsamestring=""
	end if
end function

function monthconvert(monthindex)
	select case monthindex
	case "MM1"
	monthconvert="Jan"
	case "MM2"
	monthconvert="Feb"
	case "MM3"
	monthconvert="Mar"
	case "MM4"
	monthconvert="Apr"
	case "MM5"
	monthconvert="May"
	case "MM6"
	monthconvert="Jun"
	case "MM7"
	monthconvert="Jul"
	case "MM8"
	monthconvert="Aug"
	case "MM9"
	monthconvert="Sept"
	case "MM10"
	monthconvert="Oct"
	case "MM11"
	monthconvert="Nov"
	case "MM12"
	monthconvert="Dec"
	end select
end function

function monthconvert(monthindex)
	select case monthindex
	case 1
	monthconvert="January"
	case 2
	monthconvert="Feburary"
	case 3
	monthconvert="March"
	case 4
	monthconvert="April"
	case 5
	monthconvert="May"
	case 6
	monthconvert="June"
	case 7
	monthconvert="July"
	case 8
	monthconvert="August"
	case 9
	monthconvert="September"
	case 10
	monthconvert="October"
	case 11
	monthconvert="November"
	case 12
	monthconvert="December"
	end select
end function

function shortweekdayconvert(weekdayindex)
	select case weekdayindex
	case 1
	shortweekdayconvert="Sun"
	case 2
	shortweekdayconvert="Mon"
	case 3
	shortweekdayconvert="Tue"
	case 4
	shortweekdayconvert="Wed"
	case 5
	shortweekdayconvert="Thu"
	case 6
	shortweekdayconvert="Fri"
	case 7
	shortweekdayconvert="Sat"
	end select
end function

function longweekdayconvert(weekdayindex)
	select case weekdayindex
	case 1
	longweekdayconvert="Sunday"
	case 2
	longweekdayconvert="Monday"
	case 3
	longweekdayconvert="Tuesday"
	case 4
	longweekdayconvert="Wednesday"
	case 5
	longweekdayconvert="Thursday"
	case 6
	longweekdayconvert="Friday"
	case 7
	longweekdayconvert="Saturday"
	end select
end function

function getFormApprovor(approvor)
	select case approvor
	case "1" 'job's group leader
	getFormApprovor="Group Leader"
	case "2" 'other group leader
	getFormApprovor="Other Group Leader"
	case "3" 'supervisor
	getFormApprovor="Supervisor"
	case "4" 'other supervisor
	getFormApprovor="Other Supervisor"
	case "5" 'manager
	getFormApprovor="Manager"
	case "6" 'other manager
	getFormApprovor="Other Manager"
	end select
end function

function getFormTransactor(transactor)
	select case approvor
	case "1" 'job's group leader
	getFormTransactor="Group Leader"
	case "2" 'other group leader
	getFormTransactor="Other Group Leader"
	case "3" 'supervisor
	getFormTransactor="Supervisor"
	case "4" 'other supervisor
	getFormTransactor="Other Supervisor"
	case "5" 'manager
	getFormTransactor="Manager"
	case "6" 'other manager
	getFormTransactor="Other Manager"
	end select
end function

function SubJobStatus(jobstatus,simg,aimg,text,ctext,alt,calt,apage)
	select case jobstatus
	case "0"
	simg="Opened"
	aimg="Pause,Abort,Alter"
	text="Opened"
	ctext="开放"
	alt="pause,abort,alter"
	calt="暂停,废除,变更"
	apage="PauseJob,AbortJob,AlterJob"
	case "1"
	simg="Closed"
	aimg="Repeat"
	text="Closed"
	ctext="关闭"
	alt="Repeat"
	calt="重复"
	apage="RepeatJob"
	case "2"
	simg="Paused"
	aimg="Start"
	text="Paused"
	ctext="暂停"
	alt="Start"
	calt="启动"
	apage="StartJob"
	case "3"
	simg="Locked"
	aimg="Start"
	text="Locked"
	ctext="锁定"
	alt="Unlock"
	calt="解锁"
	apage="UnlockJob"
	case "4"
	simg="Aborted"
	aimg=""
	text="Aborted"
	ctext="废除"
	alt=""
	calt=""
	apage=""
	case "5"
	simg="ShiftOut"
	aimg=""
	text="Shift Out"
	ctext="停线"
	alt=""
	calt=""
	apage=""
	end select
end function

function getApprovalLevel(level_list,check_list)


end function

'function getApproveLevelOptions(selected)
'level_name="Group Leader,Supervisor,Manager,Director,GM,VP"
'level_value=

'end function
'delete TEMP file
function DeleteFile(URL)
	set myFs=server.createObject("scripting.FileSystemObject") 
	if myFs.FileExists(URL)=true then
	myFs.DeleteFile URL 
	end if
	set myFs=nothing 
end function

function timeconvert(thistime,unit)
	if thistime<>"" then
		if unit="MM" then
		timeconvert=thistime
		elseif unit="HH" then
		timeconvert=thistime*60
		elseif unit="DD" then
		timeconvert=thistime*60*24
		else
		timeconvert=0
		end if
	else
	timeconvert=0
	end if
end function

function unitconvert(thistime,newtime)
	if thistime>0 then
		if thistime/60/24>1 and csng(thistime/60/24)=thistime/60/24 then
		unitconvert="DD"
		newtime=thistime/60/24
		elseif thistime/60>1 and csng(thistime/60)=thistime/60 then
		unitconvert="HH"
		newtime=thistime/60
		else
		unitconvert="MM"
		newtime=thistime
		end if
	else
	unitconvert=""
	newtime=0
	end if
end function

function jobworkday(fromday,today,hourtime) 
	reduceweekday=0
	if isnull(today) then
	today=date()
	end if
	interval=datediff("d",fromday,today)
	thisday=fromday
'	for ji=0 to interval
'		if weekday(thisday)=1 then
'		reduceweekday=reduceweekday+1
'		end if
'		thisday=dateadd("d",1,thisday)
'	next
	jobworkday=round(datediff("n",fromday,today)/60/24,1)-reduceweekday
	hourtime=datediff("n",fromday,today)-reduceweekday*60*24
end function
%>