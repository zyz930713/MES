<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
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
			inspectionpqc=trim(request.Form("txt_INSPECTIONPQC"))	
			Qty=trim(request.Form("txt_boxsize"))
			job_number=trim(request.Form("txt_Job_Number"))
			bind_box_id=trim(request.Form("txt_Bind_Box_Id"))
			conn.beginTrans
			
			 					
			sql="select * from job_pack_detail where box_id='"&boxid&"' and job_number='"&job_number&"'"		
				
		     	rs.open sql,conn,1,3
				if  rs.eof then
						
					rs.addnew
					rs("box_id")=boxid
					rs("job_number")=job_number	
				end if		

					rs("packed_qty")=cint(Qty)
					rs("part_number")=partnumber
					rs("customer")=customer
					rs("packed_user")=opcode
					rs("packed_time")=now()
					rs("remarks")=remarks
					rs("packed_line")=packLine
					rs("prod_shift")=prodShift
					rs("supplier")=supplier
					rs("customer_label_id")=custId
					rs("inspectionpqc")=inspectionpqc
					rs("BIND_BOX_ID")=bind_box_id
					rs.update
					rs.close				
				
		
		
			
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
			customer=trim(request.Form("txt_customer"))
			opcode=trim(request.Form("txt_opcode"))			
			BindBox_id=trim(request("txt_Bind_Box_Id"))
			partNo=trim(request.Form("txt_part"))
			if partNO="240326000001" then
			partNoA="990321100028"
			else			
			partNoA=partNo
			end if
			
			
			job_number=trim(request("txt_Job_Number"))
			Qty=trim(request.Form("txt_boxsize"))
			remarks=trim(request.Form("txt_remarks"))
			INSPECTIONPQC=trim(request("txt_INSPECTIONPQC"))			
		    boxid=trim(request.Form("txt_boxid"))
			custId=trim(request("txt_custid"))				
			isCustid=trim(request("chkCustidPack"))
			isTray=trim(request.Form("chkTrayPack"))
			isPartial=trim(request.Form("chkPartialPack"))
			packType=trim(request("txt_pack_type"))
			prodShift=trim(request.Form("txt_shift"))
					
			SQLH="SELECT WipEntityName,PRIMARYITEMDESC FROM tbl_MES_LotMaster  where WipEntityName='"&job_number&"'"			
			rsTicket.open SQLH,connTicket,1,3
			if  rsTicket.bof and rsTicket.eof then
			  member("error")="This Job ("&job_number&") does not  in system.|该工单("&job_number&")在系统中不存在，请确认工单号。"		
			else
			PRIMARYITEMDESC=rsTicket("PRIMARYITEMDESC")
			end if	
			rsTicket.close				
			packLine=mid(job_number,3,4)	
			'check partial packing
			if member("error")="" then
			   if PRIMARYITEMDESC<>partNoA then
			    member("error")="工单号与成品料号不匹配，请不要乱输入工单号与料号。"				   
			   end if			
			end if
		
				if member("error")="" then			
				'产生Customer label id
					sql="select CUSTOMER_PN,VENDOR_CODE,nvl(box_size,0) as box_size, nvl(SMALL_LABLE_SIZE,0) as SMALL_PACK,YESNO_LITTLE_LABLE,PRODUCT_TYPE from CUSTOMER_MATERIAL where item_name='"&partNo&"' and CUSTOMER_LABEL='"&CUSTOMER&"'"
					rs.open sql,conn,1,3
					if not rs.eof then
						customerPN=rs("CUSTOMER_PN")
						VENDOR_CODE=rs("VENDOR_CODE")					
						boxSize=rs("box_size")
						SMALL_PACK=rs("SMALL_PACK")
						YESNO_LITTLE_LABLE=rs("YESNO_LITTLE_LABLE")
						PRODUCT_TYPE=rs("PRODUCT_TYPE")
					else
						member("error")="This customer label("&CUSTOMER&") does not define in system.|该客户标签("&CUSTOMER&")在系统中未定义，请联系主管。"					
					end if
					rs.close
				end if		
			 '----------------只打印 BoxLable
			 IF isTray<>"Y"  and isCustid<>"Y" Then
					if member("error")="" then
						if customer<>"HuaWei" then	
						result=	PrintBoxLable		
						end if										
						if result <> "" then
							member("error")=result						
						end if	
					end if	 	
			 END IF
		'	 ----------------------打印CustmerLable
			IF isPartial<>"Y" 	and  isTray<>"Y"  then
					if member("error")="" then	
								IF isTray<>"Y" Then
									if member("error")="" then	
									   if customer<>"HuaWei" then									
											result=PrintCustmerLable
									   end if		
											if result <> "" then
												member("error")=result						
											end if
									end if	
								End If
			        end if
			
			end if
			
			 '------------------只打印小标签
			
		'	IF isPartial<>"Y"  and  isCustid<>"Y"  then			 
'							
'						 							
'									if member("error")="" then
'												
'												if YESNO_LITTLE_LABLE="YES"  then												
'													   if isnull(SMALL_PACK) or  cstr(SMALL_PACK)=0  then
'															member("error")="请联系主管配置最小包装数,否则不能打小标签"
'													   else
'															if cint(txt_count)=cint(boxSize) then
'														'	result=	PrintSMALLBoxLable(cstr(boxSize),cstr(SMALL_PACK))	
'															else
'															'result=	PrintSMALLBoxLable(cstr(txt_count),cstr(SMALL_PACK))	
'															end if													
'																								   
'													   end if											
'												end if
'															
'												if result <> "" then
'													member("error")=result						
'												end if			
'									end if
'								 
'			end if		
'			
'			'------------------打印BoxLable 和 CustmerLable
'			
'			IF isPartial="Y"  and isCustid="Y" Then
'					if member("error")="" then
'					'调用打印功能
'								result=	PrintBoxLable												
'								if result <> "" then
'									member("error")=result						
'								end if			
'					end if	 	
'	
'					
'					If customerPN="" then
'							member("error")="This part number("&partNo&") does not define customer label.|该料号("&partNo&")未定义客户标签。"
'					Else					
'							if customer<>"HuaWei" then							
'							result=PrintCustmerLable
'							end if
'							if result <> "" then
'								member("error")=result						
'							end if
'					End if
'       		END IF	
'			
'          
'	'-------------------打印箱号和小签标-----------------
'	
'	
'	IF isPartial="Y"  and  isTray="Y"  then			 
'								if member("error")="" then
'									'调用打印功能
'									result=	PrintBoxLable												
'									if result <> "" then
'										member("error")=result						
'									end if			
'								end if	 	
'						 							
'								if member("error")="" then
'											
'											if YESNO_LITTLE_LABLE="YES"  then												
'												   if isnull(SMALL_PACK) or  cstr(SMALL_PACK)=0  then
'														member("error")="请联系主管配置最小包装数,否则不能打小标签"
'												   else
'														if cint(txt_count)=cint(boxSize) then
'														result=	PrintSMALLBoxLable(cstr(boxSize),cstr(SMALL_PACK))	
'														else
'														result=	PrintSMALLBoxLable(cstr(txt_count),cstr(SMALL_PACK))	
'														end if													
'																							   
'												   end if											
'											end if
'														
'											if result <> "" then
'												member("error")=result						
'											end if			
'								end if
'								 
'	end if		



 					function PrintCustmerLable
			   
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
									countCondition=VENDOR_CODE&strYear&getOneCharMon(strMon)&strDay&strWeek&N10toC62(cint(right(packLine,1))+9,16)&prodShift								
									sql="select count_value,lm_time,count_type,count_condition from serial_count "
									sql=sql+"where count_type='"&PRODUCT_TYPE&"' and count_condition = '"&countCondition&"'"
			
								rs.open sql,conn,1,3
									if rs.eof then
										custId=countCondition&"0001"
										rs.addNew
										rs("count_type")=PRODUCT_TYPE
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
								
								if member("error")="" then
									
													
									'调用打印功能
									sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"						
									rsAsyn.open sql,conn,1,3						
										if not rsAsyn.eof then
											printName = rsAsyn("PRINTER_NAME")
										end if
									rsAsyn.close
									
									if printName="" then
										alarmMsg="Please contact engiener to set the printer for this machine.|请联系工程师为此机器设定正确的标签打印机。"
									else
									set PrintCtl=server.createobject("PrintClass.PrintCtl") 										
										DateCode=201&strYear&strMon&strDay										
										 rtnCode=PrintCtl.PrintOEMKEBCustomerLabel(printName,customer,customerPN,cstr(qty),custId,partNo)    																	
									'rtnCode="OK"
										if rtnCode<>"OK" then
											alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
										end if
								   end if
								end if	
			        PrintCustmerLable=alarmMsg
			   end function
			   
			   function PrintBoxLable			   
			  			
						'
						     if boxid="" or isnull(boxid) then
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
				
						 if member("error")="" then
								'调用打印功能
								sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"
								rsAsyn.open sql,conn,1,3
								if not rsAsyn.eof then
									printName = rsAsyn("PRINTER_NAME")
								end if
								rsAsyn.close
								if printName="" then
									alarmMsg="Please contact engiener to set the printer for this machine.|请联系工程师为此机器设定正确的标签打印机。"
								else
									set PrintCtl=server.createobject("PrintClass.PrintCtl")                  
																
										rtnCode=PrintCtl.PrintOEMKEBBoxIDLabel(printName,boxid,customer,partNo,cstr(qty),remarks,INSPECTIONPQC)
									
									'rtnCode="OK"
									if rtnCode<>"OK" then
										alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
									end if
								end if
						end if	 	
						PrintBoxLable=alarmMsg
			   end function
			   
			   
			   
			   function PrintSMALLBoxLable (sumcount,pouchpack)
					
					Traytime=int(sumcount/pouchpack)
					LastQty=sumcount-pouchpack*Traytime 
					qty=int(pouchpack)
				   '调用打印功能
					sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"						
					rsAsyn.open sql,conn,1,3						
						if not rsAsyn.eof then
						printName = rsAsyn("PRINTER_NAME")
						end if
					rsAsyn.close															
					if printName="" then
						alarmMsg="Please contact engiener to set the printer for this machine.|请联系工程师为此机器设定正确的标签打印机。"
					else
						set PrintCtl=server.createobject("PrintClass.PrintCtl")                  
						if Traytime>0 then
						for i=1 to Traytime 	
						    if   customer="HuaWei" then 
								sql="select to_char(sysdate+1,'IW') as week,to_char(sysdate,'YY')as year,to_char(sysdate,'DD') as day,to_char(sysdate,'MM') as mon from dual"
								rs.open sql,conn,1,3
									if not rs.eof then
										strYear=rs("year")
										strMon=rs("mon")
										strDay=rs("day")
										strWeek=rs("week")
									end if
								rs.close	
							      DateCode=strYear&strWeek
								                                           			 
							rtnCode=PrintCtl.PrintHWCustomerLabel(printName,customer,DateCode,customerPN,cstr(qty),LotNO,boxid,Vendor_Code)
							else														
						     rtnCode=PrintCtl.PrintKEBTrayLabel(printName,custId,cstr(qty),customerPN,partNo)
							end if
							'rtnCode="OK"
								if rtnCode<>"OK" then
									alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
								end if
						next
						end if
						if LastQty>0 then
						    if  customer="HuaWei" then 
							rtnCode=PrintCtl.PrintHWCustomerLabel(printName,customer,DateCode,customerPN,cstr(qty),LotNO,boxid,Vendor_Code)
							else
							rtnCode=PrintCtl.PrintKEBTrayLabel(printName,custId,cstr(LastQty),customerPN,partNo)
							end if
							'rtnCode="OK"
								if rtnCode<>"OK" then
									alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
								end if															
						end if
					end if
					PrintSMALLBoxLable=alarmMsg
			   end function
         
			
			if err.number <> 0 then 
						conn.rollbackTrans   '对已执行的操作回滚
					
						member("error")=err.description
					else 
						conn.commitTrans     '执行事务提交
			end if	
			member("pack_line")=packLine
			member("cust_id")=custId	
			member("INSPECTIONPQC")=INSPECTIONPQC
	        member("boxid")=boxId
			
			response.Write toJSON(member)
		case "3" 'check 2D code
		
			
	end select
	set rsAsyn=nothing 	
end if


%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Close.asp" -->