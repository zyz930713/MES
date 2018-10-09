<%
function getGroup(isshowname,showstyle,group,where,order,sequency,splitchar,gcount,action)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
gcount=0
set rsPAM=server.CreateObject("adodb.recordset")
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select 1,SG.NID,SG.GROUP_NAME,SG.GROUP_MEMBERS from SYSTEM_GROUP SG inner join FACTORY F on SG.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
gcount=rsP.recordcount
rsP.close
SQLP="Select 1,SG.NID,SG.GROUP_NAME,SG.GROUP_MEMBERS from SYSTEM_GROUP SG inner join FACTORY F on SG.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
if not rsP.eof then
	if sequency<>"" then
		aseq=split(sequency,",")
		for x=0 to ubound(aseq)
			while not rsP.eof
				if aseq(x)=rsP("NID") then
					select case showstyle
					case "OPTION"
					output=output&"<option value='"&rsP("NID")&"'"
						if rsP("NID")=group then
						output=output&"selected"
						end if
					output=output&">"&rsP("GROUP_NAME")&"</option>"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsP("GROUP_NAME")&splitchar
						else
						output=output&rsP("GROUP_NAME")
						end if
					end select
				end if
			rsP.movenext
			wend
		rsP.movefirst
		next
	else
		y=1
		while not rsP.eof
			select case showstyle
			case "OPTION"
				output=output&"<option value='"&rsP("NID")&"'"
					if rsP("NID")=group then
					output=output&"selected"
					end if
				output=output&">"&rsP("GROUP_NAME")&"</option>"
			case "TEXT"
				if y<>rcount then
				output=output&rsP("GROUP_NAME")&splitchar
				else
				output=output&rsP("GROUP_NAME")
				end if
			case "TABLE_VALUE"
				valid_material=""
				valid_machine=""
				output=output&"<tr class='t-c-GrayLight'><td><input name='group_id"&y&"' type='hidden' value='"&rsP("NID")&"'><span title='"&rsP("GROUP_MEMBERS")&"'>"&rsP("GROUP_NAME")&"</span></td>"
				SQLPAM="select MATERIAL,MACHINE from GROUP_ACTION_VALUE where GROUP_ID='"&rsP("NID")&"' and ACTION_ID='"&action&"'"
				rsPAM.open SQLPAM,conn,1,3
				if not rsPAM.eof then
				valid_material=rsPAM("MATERIAL")
				valid_machine=rsPAM("MACHINE")
				end if
				rsPAM.close
				output=output&"<td><input name='gmaterial"&y&"' type='text' id='gmaterial"&y&"' size='15' value='"&valid_material&"'></td><td><input name='gmachine"&y&"' type='text' id='gmachine"&y&"' size='15' value='"&valid_machine&"'></td></tr>"
			case "JOB_SCHEDULE"
				output=rsP("NID")
			end select
		rsP.movenext
		y=y+1
		wend
	end if
end if
getGroup=output
rsP.close
set rsPAM=nothing
set rsP=nothing
end function
%>