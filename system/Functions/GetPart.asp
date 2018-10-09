<%
function getPart(isshowname,showstyle,part,where,order,sequency,splitchar,rcount,action,defectcode,tabi,a_SUB_ID,a_SUB_Number)
	'isshowname=whether to output name of station;
	'showstyle=what's the ouput style:
	'	OPTION=select component
	'	TEXT=text words
	'where=where phrase of SQL query
	'splitchar=if ouput need to be split, what's character displayed as split
output=""
rcount=0
set rsPAM=server.CreateObject("adodb.recordset")
set rsP=server.CreateObject("adodb.recordset")
SQLP="Select 1,P.NID from Part P inner join FACTORY F on P.FACTORY_ID=F.NID "&where&order
rsP.open SQLP,conn,1,3
rcount=rsP.recordcount
rsP.close
SQLP="Select 1,P.NID,P.PART_NUMBER,P.SECTION_ID,S.SECTION_NAME from Part P inner join SECTION S on P.SECTION_ID=S.NID inner join FACTORY F on P.FACTORY_ID=F.NID"&where&order
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
						if rsP("NID")=part then
						output=output&"selected"
						end if
					output=output&">"&rsP("PART_NUMBER")&"</option>"
					case "TEXT"
						if x<>ubound(aseq) then
						output=output&rsP("PART_NUMBER")&splitchar
						else
						output=output&rsP("PART_NUMBER")
						end if
					end select
				end if
			rsP.movenext
			wend
		rsP.movefirst
		next
	else
		y=1
		tabi=1
		while not rsP.eof
			select case showstyle
			case "OPTION"
				output=output&"<option value='"&rsP("NID")&"'"
					if rsP("NID")=part then
					output=output&"selected"
					end if
				output=output&">"&rsP("PART_NUMBER")&" ("&rsP("SECTION_NAME")&")</option>"
			case "TEXT"
				if y<>rcount then
				output=output&rsP("PART_NUMBER")&splitchar
				else
				output=output&rsP("PART_NUMBER")
				end if
			case "ACTION_TABLE_VALUE"
				if section<>rsP("SECTION_ID") then
					if tabi<>1 then
					output=output&"</tbody>"
					end if
				output=output&"<tr class='t-c-greenCopy'><td height='20' colspan='3'><span style='cursor:hand' onClick='tabshow("&tabi&")'><img src='/Images/Treeimg/plus.gif' name='tabimg"&tabi&"' width='9' height='9'>&nbsp;"&rsP("SECTION_NAME")&"</span></td></tr><tbody id='tab"&tabi&"' style='display:none'>"
				tabi=tabi+1
				end if
				valid_material=""
				valid_machine=""
				output=output&"<tr class='t-c-GrayLight'><td><input name='part_id"&y&"' type='hidden' value='"&rsP("NID")&"'>"&rsP("PART_NUMBER")&"</td>"
				SQLPAM="select MATERIAL,MACHINE from PART_ACTION_VALUE where PART_ID='"&rsP("NID")&"' and ACTION_ID='"&action&"'"
				rsPAM.open SQLPAM,conn,1,3
				if not rsPAM.eof then
				valid_material=rsPAM("MATERIAL")
				valid_machine=rsPAM("MACHINE")
				end if
				rsPAM.close
				output=output&"<td><input name='material"&y&"' type='text' id='material"&y&"' size='15' value='"&valid_material&"'></td><td><input name='machine"&y&"' type='text' id='machine"&y&"' size='15' value='"&valid_machine&"'></td></tr>"	
				section=rsP("SECTION_ID")
			case "DEFECTCODE_TABLE_VALUE"
				if section<>rsP("SECTION_ID") then
					if tabi<>1 then
					output=output&"</tbody>"
					end if
				output=output&"<tr class='t-c-greenCopy'><td height='20' colspan='3'><span style='cursor:hand' onClick='tabshow("&tabi&")'><img src='/Images/Treeimg/plus.gif' name='tabimg"&tabi&"' width='9' height='9'>&nbsp;"&rsP("SECTION_NAME")&"</span></td></tr><tbody id='tab"&tabi&"' style='display:none'>"
				tabi=tabi+1
				end if
				sub_part_id=""
				output=output&"<tr class='t-c-GrayLight'><td><input name='part_id"&y&"' type='hidden' value='"&rsP("NID")&"'>"&rsP("PART_NUMBER")&"</td>"
				SQLPAM="select SUB_PART_ID from DEFECTCODE_PART_ROUTINE where DEFECTCODE_ID='"&defectcode&"' and PART_ID='"&rsP("NID")&"'"
				rsPAM.open SQLPAM,conn,1,3
				if not rsPAM.eof then
				sub_part_id=rsPAM("SUB_PART_ID")
				end if
				rsPAM.close
				output=output&"<td><select name='sub_part_id"&y&"' id='sub_part_id"&y&"'><option value=''>-- Select Type --</option>"
				for s=0 to ubound(a_SUB_ID)
					if sub_part_id=a_SUB_ID(s) then
					output=output&"<option value='"&a_SUB_ID(s)&"' selected>"&a_SUB_Number(s)&"</option>"
					else
					output=output&"<option value='"&a_SUB_ID(s)&"'>"&a_SUB_Number(s)&"</option>"
					end if
		  		next
	  			output=output&"</select></td></tr>"	
				section=rsP("SECTION_ID")
			case "JOB_SCHEDULE"
				output=rsP("NID")
			end select
		rsP.movenext
		y=y+1
		wend
	end if
end if
'if showstyle="ACTION_TABLE_VALUE" or showstyle="DEFECTCODE_TABLE_VALUE" then
'output=output&"</tbody>"
'end if
getPart=output
rsP.close
set rsPAM=nothing
set rsP=nothing
end function
%>