<%
function getStationGroupName(job_number,sheet_number,modelname)
on error resume next

set rsS=server.CreateObject("adodb.recordset")
SQLS="Select * FROM JOB where JOB_NUMBER='"&job_number&"' AND SHEET_NUMBER='"&cstr(sheet_number)&"'"
rsS.open SQLS,conn,1,3
stations_index=rsS("stations_index")


SQLS="select * from product_model where item_name='"+modelname+"'"


set rs4LevelInfo=server.CreateObject("adodb.recordset")

rs4LevelInfo.open SQLS,conn,1,3
if(rs4LevelInfo.recordcount>0) then
	family_id=rs4LevelInfo("family_id")
	series_id=rs4LevelInfo("series_group_id")
	subseries_id=rs4LevelInfo("series_id")
else
	getStationGroupName="This model:"+modelname+" is not maitained in 4 Level.<br>"+modelname+"没有维护四层结构，请联系PE."
	EXIT function
end if 

SQLS="delete from JOB_STATION_GROUP_UVS where job_number='"+job_number+"' and sheet_number='"+cstr(sheet_number)+"'"
set rsDelete1=server.CreateObject("adodb.recordset")
rsDelete1.open SQLS,conn,1,3




stations_indexArr=split(stations_index,",")


prevoius_stationGroup=""
for m=0 to ubound(stations_indexArr)

	MAPPING_ID=""
	STATION_GROUP_ID=""
	searchfamily_id=""
	searchseries_id=""
	searchsubseries_id=""
	searchmodelname=""

	if(stations_indexArr(m)<>"") then
	
			SQLS="select * from station where nid='"+stations_indexArr(m)+"'"
			set rsStation=server.CreateObject("adodb.recordset")
			rsStation.open SQLS,conn,1,3
			if(rsStation.recordcount>0) then
				station_id=rsStation("mother_station_id")
			end if 
		
			SQLS="select * from STATION_GROUP_STATION_MAPPING where FAMILY_ID='"+family_id+"' AND SERIES_ID='"+series_id+"'"
			SQLS=SQLS+" AND SUBSERIES_ID='"+subseries_id+"' AND MODEL_NAME='"+modelname+"' AND STATION_ID='"+station_id+"'"
			
			 

			set rs01=server.CreateObject("adodb.recordset")
			rs01.open SQLS,conn,1,3
			if rs01.recordcount>0 then
				for i=0 to rs01.recordcount
					MAPPING_ID=rs01("MAPPING_ID")
					STATION_GROUP_ID=rs01("STATION_GROUP_ID")
					searchfamily_id=rs01("FAMILY_ID")
					searchseries_id=rs01("SERIES_ID")
					searchsubseries_id=rs01("SUBSERIES_ID")
					searchmodelname=rs01("MODEL_NAME")
				next 
			else
				SQLS="select * from STATION_GROUP_STATION_MAPPING where FAMILY_ID='"+family_id+"' AND SERIES_ID='"+series_id+"'"
				SQLS=SQLS+" AND SUBSERIES_ID='"+subseries_id+"'AND (MODEL_NAME is null or MODEL_NAME='') AND STATION_ID='"+station_id+"'"
				'response.write SQLS &"<br>"
				
				set rs02=server.CreateObject("adodb.recordset")
				rs02.open SQLS,conn,1,3
				if rs02.recordcount>0 then
					for i=0 to rs02.recordcount
						MAPPING_ID=rs02("MAPPING_ID")
						STATION_GROUP_ID=rs02("STATION_GROUP_ID")
						searchfamily_id=rs02("FAMILY_ID")
						searchseries_id=rs02("SERIES_ID")
						searchsubseries_id=rs02("SUBSERIES_ID")
						searchmodelname=rs02("MODEL_NAME")
					next
				else
					SQLS="select * from STATION_GROUP_STATION_MAPPING where FAMILY_ID='"+family_id+"' AND SERIES_ID='"+series_id+"' AND (SUBSERIES_ID is null or SUBSERIES_ID='') AND (MODEL_NAME is null or MODEL_NAME='')"
					SQLS=SQLS+"  AND STATION_ID='"+station_id+"'"
					'response.write SQLS &"<br>"
					set rs03=server.CreateObject("adodb.recordset")
					rs03.open SQLS,conn,1,3
					if rs03.recordcount>0 then
	
						for i=0 to rs03.recordcount
							MAPPING_ID=rs03("MAPPING_ID")
							STATION_GROUP_ID=rs03("STATION_GROUP_ID")
							searchfamily_id=rs03("FAMILY_ID")
							searchseries_id=rs03("SERIES_ID")
							searchsubseries_id=rs03("SUBSERIES_ID")
							searchmodelname=rs03("MODEL_NAME")
						next
					else
						SQLS="select * from STATION_GROUP_STATION_MAPPING where FAMILY_ID='"+family_id+"'AND (SERIES_ID IS NULL OR SERIES_ID='' )"
						SQLS=SQLS+" AND (SUBSERIES_ID is null or SUBSERIES_ID='') AND (MODEL_NAME is null or MODEL_NAME='') AND STATION_ID='"+station_id+"'"
						set rs04=server.CreateObject("adodb.recordset")
						'response.write SQLS &"<br>"
						rs04.open SQLS,conn,1,3
						if rs04.recordcount>0 then
							for i=0 to rs04.recordcount
								MAPPING_ID=rs04("MAPPING_ID")
								STATION_GROUP_ID=rs04("STATION_GROUP_ID")
								searchfamily_id=rs04("FAMILY_ID")
								searchseries_id=rs04("SERIES_ID")
								searchsubseries_id=rs04("SUBSERIES_ID")
								searchmodelname=rs04("MODEL_NAME")
							next
						end if 
					end if 
				end if 
			end if 

	
	 
		SQLS="select * from STATION_GROUP_STATION_MAPPING where STATION_ID='"+station_id+"'"
		if(searchfamily_id<>"")then
			SQLS=SQLS+" and FAMILY_ID='"+searchfamily_id+"'"
		ELSE
			SQLS=SQLS+" AND (FAMILY_ID is null or FAMILY_ID='')"
		end if 
		
		if(searchseries_id<>"")then
			SQLS=SQLS+" and SERIES_ID='"+searchseries_id+"'"
		ELSE
			SQLS=SQLS+" AND (SERIES_ID is null or SERIES_ID='')"
		end if 
		

		if(searchsubseries_id<>"")then
			SQLS=SQLS+" and SUBSERIES_ID='"+searchsubseries_id+"'"
		ELSE
			SQLS=SQLS+" AND (SUBSERIES_ID is null or SUBSERIES_ID='')"
		end if 
		
		if(searchmodelname<>"")then
			SQLS=SQLS+" and MODEL_NAME='"+searchmodelname+"'"
		ELSE
			SQLS=SQLS+" AND (MODEL_NAME is null or MODEL_NAME='')"
		end if 
		
		SQLS=SQLS+" ORDER BY STATION_SEQUENCE_NO "
  	 
		
			
		set rsStation=server.CreateObject("adodb.recordset")
		rsStation.open SQLS,conn,1,3
	
	   IF(m=9) then
	   '	response.write SQLS
		'response.end
	   end if 
	 
		'response.write "STATION_GROUP_ID:" & rsStation("STATION_GROUP_ID")
		if(isnull(rsStation("STATION_GROUP_ID"))) then
			'response.write "null"
		else
			'response.write rsStation("STATION_GROUP_ID")
		end if
		
		if (prevoius_stationGroup="" or prevoius_stationGroup<>rsStation("STATION_GROUP_ID") ) then
			MAPPING_ID=rsStation("MAPPING_ID")
			UVS=""
			SQLS="SELECT UVS FROM STATION_GROUP_UVS_SETTING WHERE MAPPING_ID='"+MAPPING_ID+"' AND STATION_GROUP_ID='"+rsStation("STATION_GROUP_ID")+"'"
			
		
		
			set rsUVS=server.CreateObject("adodb.recordset")
			rsUVS.open SQLS,conn,1,3
			if(rsUVS.recordcount>0)then
				UVS=rsUVS(0)
			end if 
		
			if isnull(searchfamily_id) then
				searchfamily_id=""
			end if 
			if isnull(searchseries_id) then
				searchseries_id=""
			end if 
			if isnull(searchsubseries_id) then
				searchsubseries_id=""
			end if 
			if isnull(modelname) then
				modelname=""
			end if 
			
	
			
 
			'insert UVS
			SQLS="insert into JOB_STATION_GROUP_UVS(JOB_NUMBER,SHEET_NUMBER,STATION_GROUP_ID,STATION_ID,FAMILY_ID,SERIES_ID,SUBSERIES_ID,MODEL_NAME,UVS,SEQUENCE_NO,MAPPING_ID) values"
			SQLS=SQLS+"('"+cstr(job_number)+"','"+cstr(sheet_number)+"','"+rsStation("STATION_GROUP_ID")+"','"+stations_indexArr(m)+"','"+searchfamily_id+"','"+searchseries_id+"','"+searchsubseries_id+"','"+modelname+"',"
			SQLS=SQLS+" '"+cstr(UVS)+"',"+cstr(m)+",'"+MAPPING_ID+"')"
			
			
			set rsJOBUVS=server.CreateObject("adodb.recordset")
			rsJOBUVS.open SQLS,conn,1,3
			'insert BOM
			'SQLS="insert into JOB_STATION_GROUP_BOM SELECT '"+job_number+"','"+sheet_number+"',STATION_GROUP_ID,'"+stations_indexArr(m)+"',FAMILY_ID,SERIES_ID,SUBSERIES_ID,"
			'SQLS=SQLS+" '"+modelname+"',MATERIAL_CATEGORY_ID,MATERIAL_CATEGORY_SEQUENCE_NO,'"+cstr(m)+"','"+MAPPING_ID+"'"
			'SQLS=SQLS+" FROM STATION_GROUP_BOM_SETTING WHERE MAPPING_ID='"+MAPPING_ID+"' AND STATION_GROUP_ID='"+rsStation("STATION_GROUP_ID")+"'"
			
			'response.write SQLS &"<br>"
			'set rsJOBOM=server.CreateObject("adodb.recordset")
			'rsJOBOM.open SQLS,conn,1,3
			
			prevoius_stationGroup=rsStation("STATION_GROUP_ID")
		else
			if(prevoius_stationGroup=rsStation("STATION_GROUP_ID"))	then
				SQLS="UPDATE JOB_STATION_GROUP_UVS SET  STATION_ID=STATION_ID || ','|| '"+stations_indexArr(m)+"'"
				SQLS=SQLS+" WHERE JOB_NUMBER='"+job_number+"' AND SHEET_NUMBER='"+sheet_number+"' AND STATION_GROUP_ID='"+rsStation("STATION_GROUP_ID")+"'"
				set rsJOBUVS_Update=server.CreateObject("adodb.recordset")
				rsJOBUVS_Update.open SQLS,conn,1,3
			
				SQLS="UPDATE JOB_STATION_GROUP_BOM SET  STATION_ID=STATION_ID || ','|| '"+stations_indexArr(m)+"'"
				SQLS=SQLS+" WHERE JOB_NUMBER='"+job_number+"' AND SHEET_NUMBER='"+sheet_number+"' AND STATION_GROUP_ID='"+rsStation("STATION_GROUP_ID")+"'"
				set rsJOBUVS_Update=server.CreateObject("adodb.recordset")
				rsJOBUVS_Update.open SQLS,conn,1,3
			end if 
			prevoius_stationGroup=rsStation("STATION_GROUP_ID")		
		end if 
	end if 
next 
getStationGroupName=""
if err.number<>0 then
		getStationGroupName="Station Group设定有错误,请联系Leo Li!" & CSTR(m)
end if 
end function
%>