<%
function FunUpdateJobPart(updatetype,part_number_id,jobnumber,sheetnumber,jobtype)
set rsU=server.CreateObject("adodb.recordset")
set rsUU=server.CreateObject("adodb.recordset")
if updatetype="SINGLE" then
SQLU="select J.JOB_NUMBER,J.SHEET_NUMBER,J.JOB_TYPE,J.CURRENT_STATION_ID,J.LAST_STATION_ID,J.FINISHED_STATIONS_ID,J.STATIONS_INDEX,J.STATIONS_TRANSACTION,J.REPEATED_STATIONS_SEQUENCE,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER='"&jobnumber&"' and J.SHEET_NUMBER='"&sheetnumber&"' and J.JOB_TYPE='"&jobtype&"' and J.STATUS=0"
else
SQLU="select J.JOB_NUMBER,J.SHEET_NUMBER,J.JOB_TYPE,J.CURRENT_STATION_ID,J.LAST_STATION_ID,J.FINISHED_STATIONS_ID,J.STATIONS_INDEX,J.STATIONS_TRANSACTION,J.REPEATED_STATIONS_SEQUENCE,P.STATIONS_INDEX as P_STATIONS_INDEX,P.STATIONS_TRANSACTION as P_STATIONS_TRANSACTION from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.PART_NUMBER_ID='"&part_number_id&"' and J.STATUS=0"
end if
rsU.open SQLU,conn,1,3
if not rsU.eof then
ok=0
while not rsU.eof
	if rsU("STATIONS_INDEX")<>rsU("P_STATIONS_INDEX") then'update routine
		old_current_flag=-1
		new_current_flag=-1
		old_stations_index=""
		new_stations_index=""
		old_stations_transaction=""
		new_stations_transaction=""
		a_stations_index=split(rsU("STATIONS_INDEX"),",")
		a_p_stations_index=split(rsU("P_STATIONS_INDEX"),",")
		a_stations_transaction=split(rsU("STATIONS_TRANSACTION"),",")
		a_p_stations_transaction=split(rsU("P_STATIONS_TRANSACTION"),",")
		for i=0 to ubound(a_stations_index)'get finished stations
			if rsU("CURRENT_STATION_ID")=a_stations_index(i) then
			old_current_flag=i
			exit for
			else
			old_stations_index=old_stations_index&a_stations_index(i)&","
			old_stations_transaction=old_stations_transaction&a_stations_transaction(i)&","
			end if
		next
		for i=0 to ubound(a_p_stations_index)'get finished stations
			if rsU("CURRENT_STATION_ID")=a_p_stations_index(i) then
			new_current_flag=i
			end if
		next
		if new_current_flag<>-1 then 'no error in new stations index
			SQLUU="select * from JOB where JOB_NUMBER='"&rsU("JOB_NUMBER")&"' and SHEET_NUMBER='"&rsU("SHEET_NUMBER")&"' and JOB_TYPE='"&rsU("JOB_TYPE")&"'"
			rsUU.open SQLUU,conn,3,3
			if not rsUU.eof then
				if ubound(a_p_stations_index)<current_flag then
				rsUU("STATIONS_INDEX")=left(old_stations_index,len(old_stations_index)-1)
				rsUU("STATIONS_TRANSACTION")=left(old_stations_transaction,len(old_stations_transaction)-1)
				rsUU("LAST_STATION_ID")=rsU("CURRENT_STATION_ID")
				else
					for i=new_current_flag to ubound(a_p_stations_index)'get finished stations
						new_stations_index=new_stations_index&a_p_stations_index(i)&","
						new_stations_transaction=new_stations_transaction&a_p_stations_transaction(i)&","
					next
				rsUU("STATIONS_INDEX")=old_stations_index&left(new_stations_index,len(new_stations_index)-1)
				rsUU("STATIONS_TRANSACTION")=old_stations_transaction&left(new_stations_transaction,len(new_stations_transaction)-1)
				rsUU("LAST_STATION_ID")=a_p_stations_index(ubound(a_p_stations_index))
				end if
			ok=ok+1
			rsUU.update
			end if
			rsUU.close
		end if
	end if
rsU.movenext
wend
end if
rsU.close
set rsU=nothing
set rsUU=nothing
FunUpdateJobPart=ok
end function
%>