<%
	function GetJobStationMaterialConsume(jobnumber,sheetnumber,stationid,line_name)
		on error resume next
		'Get Part ID
		SQL="SELECT part_number_id FROM JOB WHERE JOB_NUMBER='"+jobnumber+"' and sheet_number='"+sheetnumber+"'"
		set rsJobPart=server.CreateObject("adodb.recordset")
		rsJobPart.open SQL,conn,1,3
		if(rsJobPart.recordcount>0) then
			Part_ID=rsJobPart(0)
		end if 
		
		'Get Routing ID
		routingid=""
		SQL="select mother_routing_id From part where nid='"+Part_ID+"'"
	
	
		set rsJobRouting=server.CreateObject("adodb.recordset")
		rsJobRouting.open SQL,conn,1,3
		if(rsJobRouting.recordcount>0) then
			routingid=rsJobRouting(0)
		end if 
			
		'Get Routing Type '0:normal, 1:rework,4:change model,5:slapping
		PART_TYPE=""
		SQL="select PART_TYPE From ROUTING where nid='"+routingid+"'"
		set rsJobRoutingType=server.CreateObject("adodb.recordset")
		rsJobRoutingType.open SQL,conn,1,3
		if(rsJobRoutingType.recordcount>0) then
			PART_TYPE=rsJobRoutingType(0)
		end if 
		
		'Get Station Start Qty
		JobStationStartQty=0
		SQL=" SELECT JS.station_start_quantity FROM JOB_STATIONS JS  WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+sheetnumber+"' "
		SQL=SQL+" AND JS.station_id='"+stationid+"'"
		set rsJobStartQty=server.CreateObject("adodb.recordset")
		rsJobStartQty.open SQL,conn,1,3
		if(rsJobStartQty.recordcount>0) then
			JobStationStartQty=rsJobStartQty(0)
		end if 

		'get UVS
		SQL=" SELECT * FROM JOB_STATION_GROUP_UVS WHERE job_number='"+jobnumber+"' and sheet_number='"+sheetnumber+"' and STATION_ID like '%"+stationid+"%'"
		set rsJobUVS=server.CreateObject("adodb.recordset")
		rsJobUVS.open SQL,conn,1,3
		if(rsJobUVS.recordcount>0) then
			UVS=rsJobUVS("UVS")
			STATION_GROUP_ID=rsJobUVS("STATION_GROUP_ID")
			MAPPING_ID=rsJobUVS("MAPPING_ID")
			PART_NUMBER=rsJobUVS("MODEL_NAME")
			STATION_GROUP_SEQUENCE=rsJobUVS("SEQUENCE_NO")
			STATION_INDEX=rsJobUVS("STATION_ID")
		end if 
		'end UVS
		
		'get Job Defect Qty
		SQL="  select decode(sum(defect_quantity),null,0,sum(defect_quantity)) as ScrapQty from  job_defectcodes jd, defectcode dc "
		SQL=SQL+" where jd.defect_code_id=dc.nid and jd.job_number='"+jobnumber+"' and jd.sheet_number='"+sheetnumber+"' and jd.STATION_ID = '"+stationid+"' and dc.transaction_type='2'"
		set rsDefectCodeQty=server.CreateObject("adodb.recordset")
		rsDefectCodeQty.open SQL,conn,1,3
		if(rsDefectCodeQty.recordcount>0) then
			StationDefectQty=rsDefectCodeQty(0)
		end if 
		
		'if Test Station 
		IsTestStation="0"
		FINISHED_STATIONS_ID=""
		
		
		SQL=" select * from job where job_number='"+jobnumber+"' and sheet_number='"+sheetnumber+"'"
		set rsTestStation=server.CreateObject("adodb.recordset")
		rsTestStation.open SQL,conn,1,3
		if(rsTestStation.recordcount>0) then
			StationIndex=rsTestStation("STATIONS_INDEX")
			FINISHED_STATIONS_ID=rsTestStation("FINISHED_STATIONS_ID")
		end if 
		
		'如果操作员点击了两次，忽略后一次. 2012-5-21
		arrFINISHED_STATIONS_ID=split(FINISHED_STATIONS_ID,",")
		totalsnumber=0
		for snumber=0 to ubound(arrFINISHED_STATIONS_ID)
			if(arrFINISHED_STATIONS_ID(snumber)=stationid) then
				totalsnumber=totalsnumber+1
			end if 
		next 
		if(totalsnumber>1) then
			GetJobStationMaterialConsume=""
			exit function
		end if 
		'end add
		
		
		
		arrStationIndex=split(StationIndex,",")
		FinalTestStationID=""
		for i=0 to ubound(arrStationIndex)
			if(arrStationIndex(i)<>"")then
				SQL=" select * from station where mother_station_id='SA00000171' and nid='"+arrStationIndex(i)+"'"
				set rsIsTestStation=server.CreateObject("adodb.recordset")
				rsIsTestStation.open SQL,conn,1,3
				if(rsIsTestStation.recordcount>0) then
					FinalTestStationID=rsIsTestStation("NID")
				end if 
			end if 
		next 
		if(instr(StationIndex,stationid)>=instr(StationIndex,FinalTestStationID)) then
			IsTestStation="1"
		end if 
		'end if 
		
	 
		 
		if(PART_TYPE="1" or PART_TYPE="4" or PART_TYPE="5") then
			IsTestStation="1"
		end if 
	
		
		SQL="DELETE FROM JOB_MATERIAL_CONSUME WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND STATION_ID='"+cstr(stationid)+"'"
		set rsDeleteHistory=server.CreateObject("adodb.recordset")
		rsDeleteHistory.open SQL,conn,1,3

		STATION_GROUP_HAS_BOM="0"
		SQL="select A.JOB_NUMBER,A.material_part_number,A.LINE, c.category_id,c.category_name , d.station_id"
		SQL=SQL+" From job_bom A,material_list B,material_category C ,STATION_MATERIAL_BOM D, STATION E "
		SQL=SQL+" where a.material_part_number=B.material_part_number AND b.material_type=C.category_id "
		SQL=SQL+" AND C.category_id=D.CATEGORY_ID AND d.station_id= e.mother_station_id"
		SQL=SQL+" AND A.job_number='"+jobnumber+"'"
		SQL=SQL+" AND d.routing_id='"+routingid+"'"
		SQL=SQL+" AND E.NID in('"+ replace(STATION_INDEX,",","','")+"')"
		set rsJobBOM_StationGroup=server.CreateObject("adodb.recordset")
		rsJobBOM_StationGroup.open SQL,conn,1,3
		if(rsJobBOM_StationGroup.recordcount>0)then
			STATION_GROUP_HAS_BOM="1"
		end if 
		
	
		
		SQL="select A.JOB_NUMBER,A.material_part_number,A.LINE, c.category_id,c.category_name , d.station_id"
		SQL=SQL+" From job_bom A,material_list B,material_category C ,STATION_MATERIAL_BOM D, STATION E "
		SQL=SQL+" where a.material_part_number=B.material_part_number AND b.material_type=C.category_id "
		SQL=SQL+" AND C.category_id=D.CATEGORY_ID AND d.station_id= e.mother_station_id"
		SQL=SQL+" AND A.job_number='"+jobnumber+"'"
		SQL=SQL+" AND d.routing_id='"+routingid+"'"
		SQL=SQL+" AND E.NID='"+stationid+"'"
		
		
		
		set rsJobBOM=server.CreateObject("adodb.recordset")
		rsJobBOM.open SQL,conn,1,3
		while not rsJobBOM.eof
			CategoryID=rsJobBOM("category_id")
			CategoryName=rsJobBOM("category_name")
			MaterialPartNumber=rsJobBOM("material_part_number")
			
			set rsJobMaterialRatio=server.CreateObject("adodb.recordset")
			SQL=" SELECT GET_BOM_RATIO('"+jobnumber+"','"+MaterialPartNumber+"') FROM DUAL"
			rsJobMaterialRatio.open SQL,conn,1,3
			MaterialRatio=0

			if rsJobMaterialRatio.recordcount>0 then
				MaterialRatio=rsJobMaterialRatio(0)
			end if 
			ConsumeMaterialQty=cdbl(JobStationStartQty)*cdbl(MaterialRatio)
			'calculate the job material consume			
			SQL="INSERT INTO JOB_MATERIAL_CONSUME(JOB_NUMBER,SHEET_NUMBER,STATION_ID,CATEGORY_ID,CATEGORY_NAME,MATERIAL_PART_NUMBER,JOB_STATION_START_QTY,BOM_RATIO,CONSUME_QTY,LINE_NAME,CONSUME_DATETIME)"
			SQL=SQL+" VALUES('"+jobnumber+"','"+cstr(sheetnumber)+"','"+stationid+"','"+CategoryID+"','"+CategoryName+"',"
			SQL=SQL+" '"+MaterialPartNumber+"',"+cstr(JobStationStartQty)+",'"+cstr(MaterialRatio)+"','"+cstr(ConsumeMaterialQty)+"','"+line_name+"',sysdate)"
			set rsJobMaterialConsume=server.CreateObject("adodb.recordset")
 			rsJobMaterialConsume.open SQL,conn,1,3
			
 			'get Material Unit Price
 			SQL="select decode(ITEM_COST,null,0,ITEM_COST) from product_model where item_name='"+MaterialPartNumber+"'"
			set rsUnitPrice=server.CreateObject("adodb.recordset")
 			rsUnitPrice.open SQL,conn,1,3
			if(rsUnitPrice.recordcount>0) then
				UnitPrice=rsUnitPrice(0)
			else
				rsUnitPrice=0
			end if 
			'end Unit Price
				
			'Scarp Detail
			if(IsTestStation="0") then
				SQL="INSERT INTO JOB_STATION_SCRAP_DETAIL(JOB_NUMBER,SHEET_NUMBER,PART_NUMBER,LINE_NAME,MAPPING_ID,STATION_GROUP_ID,STATION_GROUP_SEQUENCE,"
				SQL=SQL+" STATION_ID,STATION_SEQUENCE,STATION_START_QTY,STATION_SCRAP_QTY,MATERIAL_PART_NUMBER,BOM_RATIO,UNIT_PRICE,IS_UNIT_SCRAP,STATION_GROUP_UVS,STATION_GROUP_HAS_BOM)"
				SQL=SQL+" VALUES('"+jobnumber+"','"+cstr(sheetnumber)+"','"+PART_NUMBER+"','"+line_name+"','"+MAPPING_ID+"','"+STATION_GROUP_ID+"','"+cstr(STATION_GROUP_SEQUENCE)+"',"
				SQL=SQL+"'"+cstr(stationid)+"','0','"+cstr(JobStationStartQty)+"','"+cstr(StationDefectQty)+"','"+MaterialPartNumber+"','"+cstr(MaterialRatio)+"','"+cstr(UnitPrice)+"',"
				SQL=SQL+" '"+IsTestStation+"','"+cstr(UVS)+"','1')"
				set rsScrapDetail=server.CreateObject("adodb.recordset")
				rsScrapDetail.open SQL,conn,1,3
				
			end if 
			
			rsJobBOM.movenext
		wend
		
		

		if(IsTestStation="1") then
				SQL="select decode(ITEM_COST,null,0,ITEM_COST) from product_model where item_name='"+PART_NUMBER+"'"
				set rsUnitPrice_Unit=server.CreateObject("adodb.recordset")
				rsUnitPrice_Unit.open SQL,conn,1,3
				if(rsUnitPrice_Unit.recordcount>0) then
					UnitPrice_Unit=rsUnitPrice_Unit(0)
				else
					UnitPrice_Unit=0
				end if 
		
				MATERIAL_VALUE=cdbl(UnitPrice_Unit)*cdbl(StationDefectQty)	
				'SCRAP Detail Table
				SQL="INSERT INTO JOB_STATION_SCRAP_DETAIL(JOB_NUMBER,SHEET_NUMBER,PART_NUMBER,LINE_NAME,MAPPING_ID,STATION_GROUP_ID,STATION_GROUP_SEQUENCE,"
				SQL=SQL+" STATION_ID,STATION_SEQUENCE,STATION_START_QTY,STATION_SCRAP_QTY,MATERIAL_PART_NUMBER,BOM_RATIO,UNIT_PRICE,IS_UNIT_SCRAP)"
				SQL=SQL+" VALUES('"+jobnumber+"','"+cstr(sheetnumber)+"','"+PART_NUMBER+"','"+line_name+"','"+MAPPING_ID+"','"+STATION_GROUP_ID+"','"+cstr(STATION_GROUP_SEQUENCE)+"',"
				SQL=SQL+"'"+cstr(stationid)+"','0','"+cstr(JobStationStartQty)+"','"+cstr(StationDefectQty)+"','"+PART_NUMBER+"','1','"+cstr(UnitPrice_Unit)+"','1')"
				set rsScrapDetail1=server.CreateObject("adodb.recordset")
				rsScrapDetail1.open SQL,conn,1,3
				
				'SCRAP Table 
				SQL="SELECT * FROM JOB_STATION_SCRAP WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND MAPPING_ID='"+cstr(MAPPING_ID)+"' AND STATION_GROUP_ID='"+STATION_GROUP_ID+"'"
				set rsScrapIsExisting=server.CreateObject("adodb.recordset")
				rsScrapIsExisting.open SQL,conn,1,3
				if(rsScrapIsExisting.recordcount>0) then
					SQL="UPDATE JOB_STATION_SCRAP SET TOTAL_MATERIAL_VALUE=PREVIOUS_MATERIAL_VALUE+"+cstr(MATERIAL_VALUE)+",PREVIOUS_MATERIAL_VALUE=CURRENT_MATERIAL_VALUE,"
					SQL=SQL+" CURRENT_MATERIAL_VALUE="+cstr(MATERIAL_VALUE)+",SCRAP_VALUE=SCRAP_VALUE+"+cstr(MATERIAL_VALUE)+",TIMESTAMP=sysdate"
					SQL=SQL+ " WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND STATION_GROUP_ID='"+STATION_GROUP_ID+"'"
			
					set rsStationScrap1=server.CreateObject("adodb.recordset")
					rsStationScrap1.open SQL,conn,1,3
				else
					SQL="INSERT INTO JOB_STATION_SCRAP(JOB_NUMBER,SHEET_NUMBER,PART_NUMBER,LINE_NAME,MAPPING_ID,STATION_GROUP_ID,STATION_GROUP_SEQUENCE,CURRENT_MATERIAL_VALUE,"
					SQL=SQL+"PREVIOUS_MATERIAL_VALUE,TOTAL_MATERIAL_VALUE,STATION_GROUP_UVS,LABOR_RATE,OH_RATE,SCRAP_VALUE,IS_UNIT_SCRAP,TIMESTAMP,STATION_GROUP_SCRAP_QTY)"
					SQL=SQL+"VALUES('"+jobnumber+"','"+cstr(sheetnumber)+"','"+PART_NUMBER+"','"+line_name+"','"+MAPPING_ID+"','"+STATION_GROUP_ID+"','"+cstr(STATION_GROUP_SEQUENCE)+"',"
					SQL=SQL+"'"+cstr(MATERIAL_VALUE)+"','"+cstr(MATERIAL_VALUE)+"','"+cstr(MATERIAL_VALUE)+"','1','1','1','"+cstr(MATERIAL_VALUE)+"','1',sysdate,'"+cstr(StationDefectQty)+"')"
					set rsStationScrap2=server.CreateObject("adodb.recordset")
					rsStationScrap2.open SQL,conn,1,3
				end if 
		else
				'LABOR_RATE=1.0
				'OH_RATE=1.0
				SQL="SELECT * FROM LABORRATE_OH"
				set rsOH=server.CreateObject("adodb.recordset")
				rsOH.open SQL,conn,1,3
				if(rsOH.recordcount>0) then
					LABOR_RATE=cdbl(cstr(rsOH("LABORRATE")))
					OH_RATE=cdbl(cstr(rsOH("OH")))
				else
					LABOR_RATE=1.0
					OH_RATE=1.0
				end if 
				
				PreviousStationGroupTotal=0
				CurrentStationGroupTotal=0
				OverAllTotal=0
				
				'calculate the history station
				SQL="SELECT decode(SUM(BOM_RATIO*UNIT_PRICE),null,0,SUM(BOM_RATIO*UNIT_PRICE)),STATION_GROUP_UVS,STATION_GROUP_ID"
				SQL=SQL+" FROM JOB_STATION_SCRAP_DETAIL   "
				SQL=SQL+" WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND IS_UNIT_SCRAP='0' AND STATION_GROUP_ID<>'"+STATION_GROUP_ID+"' "
				SQL=SQL+" AND (STATION_GROUP_HAS_BOM='0' OR  MATERIAL_PART_NUMBER<>'-' AND STATION_GROUP_HAS_BOM='1')"
				SQL=SQL+" GROUP BY STATION_GROUP_ID,STATION_GROUP_UVS"
				set rsDetailScrap=server.CreateObject("adodb.recordset")
				rsDetailScrap.open SQL,conn,1,3
				while not rsDetailScrap.eof
					if(rsDetailScrap("STATION_GROUP_ID")<>STATION_GROUP_ID) then
							jobscrapvalue=cdbl(cstr(rsDetailScrap(0)))
							jobuvs=cdbl(cstr(rsDetailScrap(1)))
							PreviousStationGroupTotal=cdbl(PreviousStationGroupTotal)+jobscrapvalue+jobuvs*LABOR_RATE+(jobscrapvalue+jobuvs*LABOR_RATE)*OH_RATE
					end if 
					rsDetailScrap.movenext
				wend

				'if(jobnumber="9240921" and sheetnumber="4") then
					'response.write PreviousStationGroupTotal &"<br>"
					 
					'response.end
				'end if 
				
				
				'calculate the current station
				SQL="SELECT decode(SUM(BOM_RATIO*UNIT_PRICE),null,0,SUM(BOM_RATIO*UNIT_PRICE)),STATION_GROUP_UVS,STATION_GROUP_ID"
				SQL=SQL+" FROM JOB_STATION_SCRAP_DETAIL   "
				SQL=SQL+" WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND IS_UNIT_SCRAP='0' AND STATION_GROUP_ID='"+STATION_GROUP_ID+"' AND MATERIAL_PART_NUMBER<>'-'"
				SQL=SQL+" GROUP BY STATION_GROUP_ID,STATION_GROUP_UVS"
				set rsDetailScrap2=server.CreateObject("adodb.recordset")
				rsDetailScrap2.open SQL,conn,1,3
				while not rsDetailScrap2.eof
					if(rsDetailScrap2("STATION_GROUP_ID")=STATION_GROUP_ID) then
							jobscrapvalue=cdbl(cstr(rsDetailScrap2(0)))
							jobuvs=cdbl(cstr(rsDetailScrap2(1)))
							CurrentStationGroupTotal=cdbl(CurrentStationGroupTotal)+jobscrapvalue
					end if
					rsDetailScrap2.movenext
				wend
				
				'if(jobnumber="9240921" and sheetnumber="4") then
					'response.write PreviousStationGroupTotal &"<br>"
					'response.write CurrentStationGroupTotal &"<br>"
					'response.end
				'end if 
				
				SQL="SELECT * FROM JOB_STATION_SCRAP WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND MAPPING_ID='"+cstr(MAPPING_ID)+"' AND STATION_GROUP_ID='"+STATION_GROUP_ID+"'"
				set rsScrapIsExisting=server.CreateObject("adodb.recordset")
				rsScrapIsExisting.open SQL,conn,1,3
				
				PreviousStationScrapQty=0
					
				if(rsScrapIsExisting.recordcount>0) then
					CurrentStationGroupTotal=CurrentStationGroupTotal+cdbl(UVS)*LABOR_RATE+(CurrentStationGroupTotal+cdbl(UVS)*LABOR_RATE)*OH_RATE
					OverAllTotal=PreviousStationGroupTotal+cdbl(CurrentStationGroupTotal)
					STATION_INDEX="'" &replace(STATION_INDEX,",","','") &"'"
					SQL="  select decode(sum(defect_quantity),null,0,sum(defect_quantity)) as ScrapQty from job_defectcodes jd, defectcode dc "
					SQL=SQL+" where jd.defect_code_id=dc.nid and jd.job_number='"+jobnumber+"' and jd.sheet_number='"+sheetnumber+"' and jd.STATION_ID in ("+STATION_INDEX+") and dc.transaction_type='2'"
		
					set rsTotalDefectCodeQty=server.CreateObject("adodb.recordset")
					rsTotalDefectCodeQty.open SQL,conn,1,3
					if(rsTotalDefectCodeQty.recordcount>0) then
						CurrentStationTotalDefectQty=rsTotalDefectCodeQty(0)
					end if 
					
					OverAllValue=cdbl(OverAllTotal)*cdbl(CurrentStationTotalDefectQty)
					SQL="UPDATE JOB_STATION_SCRAP SET TOTAL_MATERIAL_VALUE=PREVIOUS_MATERIAL_VALUE+"+cstr(CurrentStationGroupTotal)+",PREVIOUS_MATERIAL_VALUE=CURRENT_MATERIAL_VALUE,"
					SQL=SQL+" CURRENT_MATERIAL_VALUE="+cstr(CurrentStationGroupTotal)+",SCRAP_VALUE="+cstr(OverAllValue)+","
					SQL=SQL+" TIMESTAMP=sysdate,STATION_GROUP_SCRAP_QTY='"+CSTR(CurrentStationTotalDefectQty)+"'"
					SQL=SQL+ " WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND STATION_GROUP_ID='"+STATION_GROUP_ID+"'"
					set rsStationScrap3=server.CreateObject("adodb.recordset")
					rsStationScrap3.open SQL,conn,1,3
				else
					OverAllTotal=PreviousStationGroupTotal+(CurrentStationGroupTotal+cdbl(UVS)*LABOR_RATE+(CurrentStationGroupTotal+cdbl(UVS)*LABOR_RATE)*OH_RATE)	
					OverAllValue=cdbl(OverAllTotal)*cdbl(StationDefectQty)
					
					'RESPONSE.WRITE "old:" & PreviousStationGroupTotal&"<br>"
					
					'RESPONSE.WRITE "current:" & (CurrentStationGroupTotal+cdbl(UVS)*LABOR_RATE+(CurrentStationGroupTotal+cdbl(UVS)*LABOR_RATE)*OH_RATE)&"<br>"
					
					'RESPONSE.WRITE "OverAllTotal:" & OverAllTotal &"<br>"
					
					'RESPONSE.WRITE "StationDefectQty:" & StationDefectQty &"<br>"
					
					'response.End
					
				
					SQL="INSERT INTO JOB_STATION_SCRAP(JOB_NUMBER,SHEET_NUMBER,PART_NUMBER,LINE_NAME,MAPPING_ID,STATION_GROUP_ID,STATION_GROUP_SEQUENCE,CURRENT_MATERIAL_VALUE,"
					SQL=SQL+"PREVIOUS_MATERIAL_VALUE,TOTAL_MATERIAL_VALUE,STATION_GROUP_UVS,LABOR_RATE,OH_RATE,SCRAP_VALUE,IS_UNIT_SCRAP,TIMESTAMP,STATION_GROUP_SCRAP_QTY)"
					SQL=SQL+"VALUES('"+jobnumber+"','"+cstr(sheetnumber)+"','"+PART_NUMBER+"','"+line_name+"','"+MAPPING_ID+"','"+STATION_GROUP_ID+"','"+cstr(STATION_GROUP_SEQUENCE)+"',"
					SQL=SQL+"'"+cstr(CurrentStationGroupTotal)+"','"+cstr(PreviousStationGroupTotal)+"','"+cstr(OverAllTotal)+"','"+cstr(UVS)+"','"+cstr(LABOR_RATE)+"','"+cstr(OH_RATE)+"',"
					SQL=SQL+" '"+cstr(OverAllValue)+"','0',sysdate,'"+cstr(StationDefectQty)+"')"
					set rsStationScrap4=server.CreateObject("adodb.recordset")
					rsStationScrap4.open SQL,conn,1,3
					
					if(STATION_GROUP_HAS_BOM="0")then
						SQL="SELECT * FROM JOB_STATION_SCRAP_DETAIL WHERE JOB_NUMBER='"+jobnumber+"' AND SHEET_NUMBER='"+cstr(sheetnumber)+"' AND STATION_GROUP_ID='"+STATION_GROUP_ID+"'"
						set rsStationScrapDetailExisting2=server.CreateObject("adodb.recordset")
						rsStationScrapDetailExisting2.open SQL,conn,1,3
						
							SQL="INSERT INTO JOB_STATION_SCRAP_DETAIL(JOB_NUMBER,SHEET_NUMBER,PART_NUMBER,LINE_NAME,MAPPING_ID,STATION_GROUP_ID,STATION_GROUP_SEQUENCE,"
							SQL=SQL+" STATION_ID,STATION_SEQUENCE,STATION_START_QTY,STATION_SCRAP_QTY,MATERIAL_PART_NUMBER,BOM_RATIO,UNIT_PRICE,IS_UNIT_SCRAP,STATION_GROUP_UVS,STATION_GROUP_HAS_BOM)"
							SQL=SQL+" VALUES('"+jobnumber+"','"+cstr(sheetnumber)+"','"+PART_NUMBER+"','"+line_name+"','"+MAPPING_ID+"','"+STATION_GROUP_ID+"','"+cstr(STATION_GROUP_SEQUENCE)+"',"
							SQL=SQL+"'"+cstr(stationid)+"','0','"+cstr(JobStationStartQty)+"','"+cstr(StationDefectQty)+"','-','1','0',"
							SQL=SQL+" '0','"+cstr(UVS)+"','0')"
						set rsStationScrapDetail2=server.CreateObject("adodb.recordset")
						rsStationScrapDetail2.open SQL,conn,1,3
					end if 
				end if 
		end if 		
		
		GetJobStationMaterialConsume=""
		if err.number<>0 then
				GetJobStationMaterialConsume="工单计算消耗时有错误，请联系IT！"
		end if 

	end function
%>

