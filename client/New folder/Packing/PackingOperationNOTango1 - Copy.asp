<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/PVS_Open.asp" -->
<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
On Error Resume Next
'ajax获取数据
response.Expires=0
response.CacheControl="no-cache"
'Response.Charset="utf-8"
asynid=request.QueryString("asynid")
if	len(asynid)=0 then
	asynid=request.Form("asynid")
end if
if len(asynid)>0 then
	set rsAsyn=server.CreateObject("adodb.recordset")
	
	dim member
	set member=jsObject()
	select case asynid
	
		case "1"'Package Save			
			opcode=trim(request.Form("txt_opcode"))
			boxid=trim(request.Form("txt_boxid"))
			customer=trim(request.Form("txt_customer"))
			remarks=trim(request.Form("txt_remarks"))
			packType=trim(request("txt_pack_type"))
			packLine=trim(request.Form("txt_pack_line"))
			partnumber=trim(request.Form("txt_part"))
			prodShift=trim(request.Form("txt_shift"))
			supplier=trim(request.Form("txt_supplier"))
			custId=trim(request.Form("txt_custid"))		
			INSPECTIONPQC=trim(request.Form("txt_INSPECTIONPQC"))	
			conn.beginTrans
			
			for i=1 to request.Form("txt_jobno").count
				'更新Job的Good Qty=Good Qty - Pack Qty，Packed Qty = Packed Qty + Pack Qty
				ipackqty=request.Form("txt_packqty")(i)
				ijobno=request.Form("txt_jobno")(i)
			
			
			
			
			
				if len(trim(ijobno))>0 then
			
				ajobnumber=split(ijobno,"-")
				strSheetNo=ajobnumber(ubound(ajobnumber))
				if isnumeric(strSheetNo) then
					sheetnumber = cint(strSheetNo)
				end if
				jobnumber=left(ijobno,len(ijobno)-4)
			
			
				existPackQty=0
					
					sql="select * from job_pack_detail where box_id='"&boxid&"' and job_number='"&jobnumber&"' and sheet_number='"&sheetnumber&"'"
					rs.open sql,conn,1,3
				if not rs.eof then
						existPackQty=csng(rs("packed_qty"))
						
				else
						rs.addnew
						rs("box_id")=boxid
						rs("job_number")=jobnumber
						rs("sheet_number")=sheetnumber
						conn.execute("update job set JOB_GOOD_QTY=JOB_GOOD_QTY-'"+ipackqty+"'  where job_number='"&jobnumber&"' and sheet_number='"&sheetnumber&"'")		
                end if
				
					rs("packed_qty")=ipackqty
					rs("part_number")=partnumber
					rs("customer")=customer
					rs("packed_user")=opcode
					rs("packed_time")=now()
					rs("remarks")=remarks
					rs("packed_line")=packLine
					rs("prod_shift")=prodShift
					rs("supplier")=supplier
					rs("INSPECTIONPQC")=INSPECTIONPQC
					rs.update
					rs.close		
				
			
			  end if
			  
			
					
						
			next
			
			
			if len(trim(boxId))>0 then
            sql="update job_pack_detail set customer='"&customer&"',remarks='"&remarks&"',customer_label_id='"&custId&"' where box_id ='"&boxId&"'"
			conn.execute(sql)					
		    end if
			
			
			
				member("error")=sqlA
			
			if err.number <> 0 then 
				conn.rollbackTrans   '对已执行的操作回滚
				member("error")="Save failed.保存失败.|"&err.description
			
				response.Write toJSON(member)
			else 
			    conn.commitTrans     '执行事务提交
			    member("message")="Save succesful.|保存成功."
				response.Write toJSON(member)
			end if	
		case "2"'label print
			opcode=trim(request.Form("txt_opcode"))
			boxid=trim(request.Form("txt_boxid"))
			customer=trim(request.Form("txt_customer"))
			remarks=trim(request.Form("txt_remarks"))
			boxSize=trim(request.Form("txt_boxsize"))
			isPartial=trim(request.Form("chkPartialPack"))
			isTray=trim(request.Form("chkTrayPack"))
			packType=trim(request("txt_pack_type"))
			partNo=trim(request.Form("txt_part"))
			packLine=trim(request.Form("txt_pack_line"))
			txt_count=trim(request("txt_count"))
			custId=trim(request("txt_custid"))			
			isBoxid=trim(request("chkBoxidPack"))
			isCustid=trim(request("chkCustidPack"))
			prodShift=trim(request("txt_shift"))
			INSPECTIONPQC=trim(request("txt_INSPECTIONPQC"))
			jobno=""			
			qty=0			
			for i=1 to request.Form("txt_packqty").count
				if len(trim(request.Form("txt_packqty")(i)))>0 then
					qty=qty+cint(request.Form("txt_packqty")(i))
					jobno=jobno&trim(request.Form("txt_jobno")(i))&","
				end if				
			next
			if len(jobno)>0 then
				jobno=left(jobno,len(jobno)-1)
			end if						
			member("boxid")=boxid
			'check partial packing
			if boxSize="" then
				boxSize="0"
			end if
		
			if member("error")="" then
				'调用打印功能
				sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"
				rsAsyn.open sql,conn,1,3
				if not rsAsyn.eof then
					printName = rsAsyn("PRINTER_NAME")
				end if
				rsAsyn.close
				if printName="" then
					member("error")="Please contact engiener to set the printer for this machine.|请联系工程师为此机器设定正确的标签打印机。"
				else
					set PrintCtl=server.createobject("PrintClass.PrintCtl")                  
					rtnCode=PrintCtl.PrintOEMKEBBoxIDLabel(printName,boxid,customer,partNo,cstr(qty),remarks,INSPECTIONPQC)
					'rtnCode="OK"
					if rtnCode<>"OK" then
						member("error")="Label print error.打印标签发生错误.|" & rtnCode		
					end if
				end if
			end if	 	

		customerPN=""
		customerDef=""
		customerLabel=""
'产生Customer label id
		sql="select customer_pn,customer_define,customer_label from product_model where item_name='"&partNo&"'"
		rs.open sql,conn,1,3
		if not rs.eof then
		customerPN=rs("customer_pn")
		customerDef=rs("customer_define")
		customerLabel=rs("customer_label")
		end if
		rs.close
	if customerLabel="" then
	member("error")="This part number("&customerPN&") does not define customer label.|该料号("&customerPN&")未定义客户标签。"
	else
		sql="select to_char(sysdate+1,'IW') as week,to_char(sysdate,'Y')as year,to_char(sysdate,'DD') as day,to_char(sysdate,'MM') as mon from dual"
		rs.open sql,conn,1,3
				if not rs.eof then
				strYear=rs("year")
				strMon=rs("mon")
				strDay=rs("day")
				strWeek=rs("week")
				end if
		rs.close
		conn.beginTrans
		
		
			if custId="" then
			countCondition=customerDef&strYear&getOneCharMon(strMon)&strDay&strWeek&N10toC62(cint(right(packLine,1))+9,16)&prodShift
			
					sql="select count_value,lm_time,count_type,count_condition from serial_count "
					sql=sql+"where count_type='"&customerLabel&"' and count_condition = '"&countCondition&"'"
					
					rs.open sql,conn,1,3
					if rs.eof then
					custId=countCondition&"0001"
					rs.addNew
					rs("count_type")=customerLabel
					rs("count_condition")=countCondition
					rs("count_value")=1
					rs("lm_time")=now()
					else
					rs("count_value")=clng(rs("count_value"))+1
					custId=countCondition&repeatstring(rs("count_value"),"0",4)
					rs("lm_time")=now()
					end if
					rs.update
					rs.close
			end if
'调用打印功能
		sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"				

		
		rsAsyn.open sql,conn,1,3						
			if not rsAsyn.eof then
			printName = rsAsyn("PRINTER_NAME")
			end if
		rsAsyn.close

		if printName="" then
		member("error")="Please contact engiener to set the printer for this machine.|请联系工程师为此机器设定正确的标签打印机。"
		else
		set PrintCtl=server.createobject("PrintClass.PrintCtl")     
		      
	    rtnCode=PrintCtl.PrintOEMKEBCustomerLabel(printName,customer,customerPN,cstr(qty),custId,partNo)
		'rtnCode="OK"
			if rtnCode<>"OK" then
			member("error")="Label print error.打印标签发生错误.|" & rtnCode		
			end if
		end if

 end if


	
			
			
            
			
		

			if err.number <> 0 then 
						conn.rollbackTrans   '对已执行的操作回滚
						member("error")=err.description
					else 
						conn.commitTrans     '执行事务提交
			end if	
				
			member("cust_id")=custId	
			member("INSPECTIONPQC")=INSPECTIONPQC
	
			response.Write toJSON(member)
		case "3" 'check 2D code
		
			opcode=trim(request.Form("txt_opcode"))
			inputJob=trim(request("txt_2DCode"))			
			boxId=trim(request("txt_boxid"))
			packType=trim(request.Form("txt_pack_type"))
			partNo=trim(request.Form("txt_part"))
			packLine=trim(request.Form("txt_pack_line"))
			boxSize=trim(request.Form("txt_boxsize"))
			prodShift=trim(request.Form("txt_shift"))
			supplier=trim(request.Form("txt_supplier"))
			codeLen=len(codeValue)
			Pcount=trim(request("txt_count"))
			
			
			if instr(inputJob,"-")=0 then
			member("error")="Job Number is error, please scan sub job！<br>工单错误，请扫描子工单！"
			end if
			
			if member("error")="" then
				ajobnumber=split(inputJob,"-")
				strSheetNo=ajobnumber(ubound(ajobnumber))
	
				if len(strSheetNo)>3 then
					member("error")="Job Number is error, please scan sub job！<br>细分工单单编号超过3位错误，请扫描子工单！"
				end if

			end if	
			
			if member("error")="" then	
				if isnumeric(strSheetNo) then
					sheetnumber = cint(strSheetNo)
				else
					member("error")="Job Number is error, please re-scan sub job！<br>工单错误，请重新扫描子工单！"
			    end if
			    jobnumber=left(inputJob,len(inputJob)-4)
			end if	
			
			
			   IF member("error")="" then	
			
				sql="select * from job where job_number='"&jobnumber&"' and  sheet_number='"&sheetnumber&"'"			
'				
				rsAsyn.open sql,conn,1,3
			
					If rsAsyn.eof then
						member("error")="This JOB Number("&inputJob&") dose not exist.|此工单不存在("&inputJob&")不存在或没有开站不能直接打包"
                   
				    elseif   clng(rsAsyn("JOB_GOOD_QTY"))<=0 or isnull(rsAsyn("JOB_GOOD_QTY")) then
						   	member("error")="此工单可包装良品数量为零，不能打包"
					else
						if member("error")="" and packType="FG"  then
						
								job_number=rsAsyn("job_number")
								sql="select job_number from job_master_store_pre where job_number='"&rsAsyn("job_number")&"' and instr(sub_job_numbers,'"&inputJob&"')>0 "
								rs.open sql,conn,1,3
								if rs.eof then
								member("error")="This Job Number("&inputJob&") has not done pre store .|扫描的工单("&inputJob&")未做预入库."
								end if
								rs.close
						end if	
						
						
						if member("error")="" and packType="FG" or PackType="SCRAP" then
							if partNo = "" then
								partNo=rsAsyn("part_number_tag")
							elseif partNo <> rsAsyn("part_number_tag") then
								member("error")="This Job_Number("&inputJob&") is mix packing part number.|该二维码的料号("&rsAsyn("part_number_tag")&")与已包装的料号不一致."					
							end if
							
							if packLine = "" then
								packLine=rsAsyn("line_name")
							elseif packLine<>rsAsyn("line_name") then
								member("error")="This Job_Number("&inputJob&") is mix packing line.|该二维码的线别("&rsAsyn("line_name")&")与已包装的线别不一致."
							end if
					
											
							if boxSize="" then
								sql="select nvl(box_size,0) as box_size  from product_model where item_name='"&partNo&"'"
								boxSize=0
								rs.open sql,conn,1,3
								if not rs.eof then
									boxSize=cint(rs("box_size"))
								end if
								rs.close
							end if
						end if
						
						if member("error")="" and packType="FG"  then
						'PCount=0
						
						
								if cint(rsAsyn("JOB_GOOD_QTY")) < boxSize then
									JOB_GOOD_QTY =rsAsyn("JOB_GOOD_QTY")					
								else
									JOB_GOOD_QTY =boxSize						
								end if
								
								if Pcount="" or isnull(Pcount) then						
									PCount=JOB_GOOD_QTY
								else
									PCount=PCount+cint(JOB_GOOD_QTY)
								end if
								
								if cint(PCount)>cint(boxSize) then
									 JOB_GOOD_QTY=cint(JOB_GOOD_QTY)-(cint(PCount)-cint(boxSize))	
									 PCount= cint(PCount)-(cint(rsAsyn("JOB_GOOD_QTY"))-cint(JOB_GOOD_QTY))					' 例： 一箱360 所有工单良品相加不能大于360
								end if
								member("job_number")=inputJob
								member("part_number")=rsAsyn("part_number_tag")
								member("pack_line")=rsAsyn("LINE_NAME")						
								member("boxsize")=boxSize						
						        member("packqty")=JOB_GOOD_QTY						     
								member("JOB_GOOD_QTY")=cint(rsAsyn("JOB_GOOD_QTY"))-cint(JOB_GOOD_QTY)  '良品可包装数量
								member("PCount")=PCount
								member("prod_shift")=rsAsyn("SHIFT")
								member("supplier")=rsAsyn("SUPPLIER")	
								if boxId="" or isnull(boxId) then
									'生成box id						
									countType="BoxId"
									countCondition=packType&packLine&formatdate(Now,"ymmddhh")
									sql="select count_value,lm_time,count_type,count_condition from serial_count "
									sql=sql+"where count_type='"&countType&"' and count_condition = '"&countCondition&"'"
									rs.open sql,conn,1,3
									if rs.eof then
										boxId=countCondition&"001"
										rs.addNew
										rs("count_type")=countType
										rs("count_condition")=countCondition
										rs("count_value")=1
										rs("lm_time")=now()
									else
										rs("count_value")=clng(rs("count_value"))+1
										boxId=countCondition&repeatstring(rs("count_value"),"0",3)
										
										rs("lm_time")=now()
									end if
									rs.update
									rs.close													
								end if
								member("boxid")=boxId						
						end if
					
					
					End if
					
				rsAsyn.close	
			   End IF 
'							'if member("error")="" and packType="FG"  then
'	'							job_number=rsAsyn("job_number")
'	'							sql="select job_number from job_master_store_pre where job_number='"&rsAsyn("job_number")&"' and instr(sub_job_numbers,'"&inputJob&"')>0 "
'	'							rs.open sql,conn,1,3
'	'							if rs.eof then
'	'							member("error")="This Job Number("&inputJob&") has not done pre store .|扫描的工单("&inputJob&")未做预入库."
'	'							end if
'	'							rs.close
'	'						end if	
'	'				
'	'				  
'	'					
'	'					if member("error")="" and packType="FG" or PackType="SCRAP" then
'	'						if partNo = "" then
'	'							partNo=rsAsyn("part_number_tag")
'	'						elseif partNo <> rsAsyn("part_number_tag") then
'	'							member("error")="This 2D Code("&codeValue&") is mix packing part number.|该二维码的料号("&rsAsyn("part_number_tag")&")与已包装的料号不一致."					
'	'						end if
'	'						
'	'						if packLine = "" then
'	'							packLine=rsAsyn("line_name")
'	'						elseif packLine<>rsAsyn("line_name") then
'	'							member("error")="This 2D Code("&codeValue&") is mix packing line.|该二维码的线别("&rsAsyn("line_name")&")与已包装的线别不一致."
'	'						end if
'	'				
'	'										
'	'						if boxSize="" then
'	'							sql="select nvl(box_size,0) as box_size  from product_model where item_name='"&partNo&"'"
'	'							boxSize=0
'	'							rs.open sql,conn,1,3
'	'							if not rs.eof then
'	'								boxSize=cint(rs("box_size"))
'	'							end if
'	'							rs.close
'	'						end if
'	'						'检查Pack Qty是否超过Job 的Good Qty
'	'						
'	'						'if member("error")="" and packType="FG"  then
'	''	
'	''						qty=0			
'	''						for i=1 to request.Form("txt_packqty").count
'	''							if len(trim(request.Form("txt_packqty")(i)))>0 then
'	''								ipackqty=trim(request.Form("txt_packqty")(i))
'	''								ijobno=trim(request.Form("txt_jobno")(i))
'	''								'if ijobno=rsAsyn("job_number") and cint(ipackqty)>(goodQty-1) then
'	''									if goodQty = 0 then
'	''									member("error")="The pack qty("&ipackqty&") of "&ijobno&" exceeds good qty("&goodQty&")|"&ijobno&"的包装数量("&ipackqty&")超过良品数量("&goodQty&")"
'	''									response.Write toJSON(member)
'	''									response.End()
'	''									end if
'	''								qty=qty+cint(request.Form("txt_packqty")(i))
'	''							end if
'	''						next
'	''						
'	''						end if
'	'						
'	'								'if (qty+1) >boxSize then
'	''								member("error")="Total pack qty("&qty&") exceeds box qty("&boxSize&").|包装数量("&qty&")超过纸箱可装的数量("&boxSize&")."
'	''								end if
'	''								set rs1=server.CreateObject("adodb.recordset")
'	''								sqlc="select COUNT(*) P_Count from job_2d_code where box_id='"&boxId&"'" 
'	''								rs1.open sqlc,conn,1,3
'	''								if  rs1("P_Count") <> "0" then
'	''								P_Count=rs1("P_Count")
'	''								end if
'	''							   if cint(P_Count) >= cint(boxSize) then
'	''							   member("error")="包装数量超过纸箱可装的数量("&boxSize&")."
'	''							   end if
'	'					
'	'					   end if
'						
'						

'				
				if err.number <> 0 then 
					conn.rollbackTrans   '对已执行的操作回滚
					member("error")=err.description
				else 
					conn.commitTrans     '执行事务提交			    				
				end if
		
			response.Write toJSON(member)
	end select
	set rsAsyn=nothing 	
end if


%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/PVS_Close.asp" -->