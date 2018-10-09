<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/Functions/PartCheck.asp" -->
<%
On Error Resume Next
'ajax获取数据
response.Expires=0
response.CacheControl="no-cache"
errMsg=""
asynid=request.QueryString("asynid")
if	len(asynid)=0 then
	asynid=request.Form("asynid")
end if
if len(asynid)>0 then
	dim member
	set member=jsObject()
	select case asynid
		case "1" 'input box id	
			txt_RewNumber=request("txt_RewNumber")
			jobNumber=request("txt_job")
			totalQty=request("hidTotalQty")
			qty=0
			'sql="select sum(packed_qty) qty from job_pack_detail where box_id='"&boxId&"'"			
			'rs.open sql,conn,1,3
			'if not rs.eof then
				'qty=rs("qty")
				'member("boxid")=boxId
				member("qty")=txt_RewNumber				
			'end if
			'rs.close
			'if isnull(qty) then
			'	errMsg="This box id("&boxId&") dose not exist.|该箱号("&boxId&")不存在."
			'end if	
			
			'sql="select * from job_pack_detail where box_id='"&boxId&"'"			
			'rs.open sql,conn,1,3
			'if not rs.eof then
			'	Packed_Line=rs("Packed_Line")
			'end if
			'rs.close

			
			
			
			if errMsg="" then
				sql="select * from job_master where job_number='"&jobNumber&"'"
				rs.open sql,conn,1,3
				if rs.eof then			
					sql="SELECT ORGANIZATIONID,WIPENTITYNAME,PRIMARYITEMDESC,LINECODE,STARTQUANTITY,U.CODE CREATE_CODE,U.NAME CREATE_NAME,DATECREATION,CREATEPERSON,DATELASTUPDATE,LASTUPDATEPERSON,BOM_LABEL FROM TBL_MES_LOTMASTER M LEFT JOIN TBL_USER_ID U ON M.CREATEPERSON=U.ID WHERE WIPENTITYNAME='"&jobNumber&"'"
					rsTicket.open sql,connTicket,1,3					
					if rsTicket.eof then
						errMsg="This job number("&jobNumber&") dose not exist.|该工单("&jobNumber&")不存在."
					else
						partNo=rsTicket("PRIMARYITEMDESC")
						lineName=rsTicket("LINECODE")
						set rsL=server.createobject("adodb.recordset")
						sql="select nid from line where status=1 and line_name='"&lineName&"'"
						rsL.open sql,conn,1,3
						if rsL.eof then
							errMsg="Line Name("&linename&") of this job does not existed.|此工单的线别("&linename&")不存在"
						else						
							facId=""
							partId=""
							sql="select nid,factory_id,part_rule from part order by meet_priority, part_rule desc"						
							set rsP=server.createobject("adodb.recordset")
							rsP.open sql,conn,1,3
							while not rsP.eof 
								aryPartRule=split(rsP("part_rule"),",")
								for p=0 to ubound(aryPartRule)
									if PartCheck(trim(aryPartRule(p)),partNo) then
										partId=rsP("nid")
										facId=rsP("factory_id")
										exit for
									end if
								next
								rsP.movenext
							wend
							rsP.close
							set rsP=nothing
							if facId <> "" then
								rs.addnew
								rs("ORGANIZATION_ID")=rsTicket("ORGANIZATIONID")
								rs("JOB_NUMBER")=jobNumber
								rs("PART_NUMBER_ID")=partId
								rs("PART_NUMBER_TAG")=partNo
								rs("FACTORY_ID")=facId
								rs("LINE_NAME")=lineName
								rs("ASSEMBLY_INPUT_QUANTITY")=clng(rsTicket("STARTQUANTITY"))
								rs("START_QUANTITY")=clng(rsTicket("STARTQUANTITY"))
								rs("REWORK_GOOD_QUANTITY")="0"
								rs("INPUT_TIME")=now()
								rs("CREATE_CODE")=rsTicket("CREATE_CODE")
								rs("CREATE_NAME")=rsTicket("CREATE_NAME")
								rs("ERP_CREATE_TIME")=rsTicket("DATECREATION")
								rs("ERP_CREATE_BY")=rsTicket("CREATEPERSON")
								rs("ERP_LAST_UPDATE_TIME")=rsTicket("DATELASTUPDATE")
								rs("ERP_LAST_UPDATE_BY")=rsTicket("LASTUPDATEPERSON")	
								rs("BOM_LABEL")=rsTicket("BOM_LABEL")
								rs.update
							else
								errMsg="This part number("&partNo&") does not set routing.|此料号("&partNo&")没有设定制程。"
							end if
						end if
						rsL.close
						set rsL=nothing
					end if
					rsTicket.close
					set rsTicket=nothing					
				end if
				if errMsg="" then
					remainQty=csng(rs("START_QUANTITY"))-csng(rs("REWORK_GOOD_QUANTITY"))					
					if csng(totalQty)+csng(qty)>remainQty then
						errMsg="Job remain qty is not enough.|工单剩余数量不足。"
					end if					
				end if
				
				'linename=rs("LINE_NAME")
				'if Packed_Line<>rs("LINE_NAME") then
			    'errMsg="箱号线别与工单的线别("&linename&")不一样！"
		        'end if				
                rs.close
			
			end if
			
			if errMsg<>"" then
				member("error")=errMsg
			end if
			response.Write toJSON(member)		
		case "2" 'print
			opCode=trim(request("txt_opcode"))
			startStation=trim(request.Form("sltStartStation"))
			job=trim(request.Form("txt_job"))
			totalQty=request.Form("hidTotalQty")
			
			conn.begintrans
			sql="select start_quantity, rework_good_quantity,part_number_tag,part_number_id,line_name,assembly_good_quantity,assembly_yield,final_good_quantity,final_input_quantity,final_yield,last_update_time,confirm_good_quantity,first_store_time,store_status,complete_date,average_cycle_time,input_time,defectcode_quantity from job_master where job_number='"&job&"'"
			rs.open sql,conn,1,3
			if not rs.eof then
				reworkQty=csng(totalQty)+csng(rs("rework_good_quantity"))
				if reworkQty>csng(rs("start_quantity")) then
					errMsg="Job remain qty is not enough.|工单剩余数量不足。"
				else
					model=rs("part_number_tag")
					partId=rs("part_number_id")
					lineName=rs("line_name")
					rs("rework_good_quantity")=reworkQty
					rs("last_update_time")=now()
					if startStation="Packing" then
						rs("assembly_good_quantity")=csng(rs("assembly_good_quantity"))+csng(totalQty)
						rs("assembly_yield")=csng(rs("assembly_good_quantity"))/(csng(rs("defectcode_quantity"))+csng(rs("assembly_good_quantity")))
						rs("final_good_quantity")=csng(rs("final_good_quantity"))+csng(totalQty)
						rs("final_input_quantity")=csng(rs("final_input_quantity"))+csng(totalQty)
						rs("final_yield")=csng(rs("final_good_quantity"))/csng(rs("final_input_quantity"))
						rs("confirm_good_quantity")=csng(rs("confirm_good_quantity"))+csng(totalQty)
						if isnull(rs("first_store_time")) then
							rs("first_store_time")=now()
						end if
						if csng(rs("confirm_good_quantity"))=csng(rs("start_quantity")) then
							rs("store_status")="1"
							rs("complete_date")=now()
							rs("average_cycle_time")=cstr(datediff("n",rs("input_time"),now()))
						end if						
					end if					
					rs.update
				end if
			end if
			rs.close
			if errMsg="" then				
				'get start station id
				sql="select a.nid from station a,part b where a.routing_id=b.mother_routing_id and b.nid='"&partId&"' and a.mother_station_id='"&startStation&"'"
				rs.open sql,conn,1,3
				if not rs.eof then
					startStation=rs("nid")
				end if
				rs.close
				
				'general sub job
				sheetNo=1
				SQL="select * from rework_printing where job_number = '"&job&"' order by seq desc "
				rs.open SQL,conn,1,3
				if not rs.eof then
					sheetNo=cint(rs("seq"))+1
				end if
				subjob=job&"-"&repeatstring(sheetNo,"0",3)
				rs.addnew
				rs("job_number")=job
				rs("qty")=totalQty
				rs("batchno")=subjob
				rs("station_name")=startStation
				rs("seq")=sheetNo
				rs("model")=model
				rs.update
				rs.close
				
				'get box id list
				boxIdList=""
				for i=1 to request.Form("txt_boxid").count
					boxId=request.Form("txt_boxid")(i)
					if len(trim(boxId))>0 then
						boxIdList=boxIdList&boxId&"','"
					end if
				next
				boxIdList=left(boxIdList,len(boxIdList)-3)
				
				if startStation="Packing" then 'do packing
					sql="select factory_id,supplier from job where (job_number,sheet_number) in (select job_number,sheet_number from job_2d_code where box_id='"&boxId&"') "
					rs.open sql,conn,1,3
					if not rs.eof then
						factoryId=rs("factory_id")
						supplier=rs("supplier")
					end if
					rs.close
					
					'insert job 
					sql="select * from job where job_number='"&job&"' and sheet_number='"&sheetNo&"'"
					rs.open sql,conn,1,3
					    if rs.eof then
						rs.addnew
						rs("job_number")=job
						rs("sheet_number")=sheetNo
						rs("part_number_id")=partId
						rs("start_time")=now()
						rs("close_time")=now()
						rs("current_station_id")=startStation
						rs("finished_stations_id")=startStation
						rs("first_station_id")=startStation
						rs("last_station_id")=startStation
						rs("status")="1"
						rs("previous_station_close_time")=now()
						rs("previous_station_id")=startStation
						rs("job_start_quantity")=totalQty
						rs("job_good_quantity")=totalQty
						rs("stations_index")=startStation
						rs("stations_transaction")="0"
						rs("stations_routine")="0"
						rs("opened_stations_id")=startStation
						rs("repeated_stations_sequence")="1"
						rs("line_name")=lineName
						rs("part_number_tag")=model
						rs("factory_id")=factoryId
						rs("cycle_time")="0"
						rs("job_type")="R"
						rs("job_assembly_yield")="1"
						rs("job_priority")="4"
						rs("new_job_type")="REWORK"
						rs("supplier")=supplier
						rs.update						
					end if
					rs.close
					
					'insert job_inventory
					sql="select job_number,part_number,good_qty,SCRAP_QTY,lm_user,lm_time from job_inventory where job_number='"&job&"'"
					rs.open sql,conn,1,3
					if rs.eof then
						rs.addnew
						rs("job_number")=job
						rs("part_number")=model
						rs("good_qty")=totalQty
						rs("SCRAP_QTY")=0
					else
						rs("good_qty")=csng(rs("good_qty"))+csng(totalQty)
					end if
					rs("lm_user")=opCode
					rs("lm_time")=now()
					rs.update
					rs.close
					
				end if
				
				'move job_2d_code to history
				'sql="INSERT INTO JOB_2D_CODE_HIST(CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER)"
'				sql=sql+" SELECT CODE,JOB_NUMBER,SHEET_NUMBER,LM_USER,LM_TIME,DEFECT_CODE_ID,BOX_ID,IS_PACKED,PACK_TIME,PACK_USER FROM JOB_2D_CODE "
'				sql=sql+" WHERE BOX_ID IN ('"&boxIdList&"')"
'				rs.open sql,conn,1,1
'				
'				'update job_2d_code for new job
'				sql="update job_2d_code set job_number='"&job&"', sheet_number='"&sheetNo&"',lm_user='"&opCode&"',lm_time=sysdate,defect_code_id='',"
'				sql=sql+"box_id='',is_packed='',pack_time='',pack_user='' where box_id in ('"&boxIdList&"')"
'				rs.open sql,conn,1,1				
'				
'				'move job_pack_detail to history
'				sql="INSERT INTO JOB_PACK_DETAIL_HIST(BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED"
'				sql=sql+",PALLET_ID,SHIPPED_USER,SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,DELIVERY_ID,WHREC_USER,WHREC_TIME)"
'				sql=sql+" SELECT BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,IS_SHIPPED,PALLET_ID,SHIPPED_USER,"
'				sql=sql+"SHIPPED_TIME,CUSTOMER_LABEL_ID,PACKED_LINE,PROD_SHIFT,SUPPLIER,PLAN_ID,STACK_USER,STACK_TIME,DELIVERY_ID,WHREC_USER,WHREC_TIME FROM JOB_PACK_DETAIL "
'				sql=sql+" WHERE BOX_ID IN ('"&boxIdList&"')"
'				rs.open sql,conn,1,1
'				
'				'delete job_pack_detail
'				sql="delete from job_pack_detail where box_id in ('"&boxIdList&"')"
'				rs.open sql,conn,1,1								
				
				member("subjob")=subjob
			end if
			
			if err.number <> 0 then 
				conn.rollbackTrans   '对已执行的操作回滚
				errMsg=errMsg+"Save failed.保存失败.|"&err.description
			else 
			    conn.commitTrans     '执行事务提交
			    'member("message")="Save succesful.|保存成功."
			end if	
			if errMsg<>"" then
				member("error")=errMsg
			end if
			response.Write toJSON(member)
		end select	
end if
%>
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->