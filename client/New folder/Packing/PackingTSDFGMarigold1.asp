<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/TSD_OpenMarigold.asp" -->
<!--#include virtual="/Components/JSON_2.0.4.asp" -->
<!--#include virtual="/Components/JSON_UTIL_0.1.1.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<%
On Error Resume Next
'ajax��ȡ����
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
			
			conn.beginTrans
			
			for i=1 to request.Form("txt_jobno").count
				'����Job��Good Qty=Good Qty - Pack Qty��Packed Qty = Packed Qty + Pack Qty
				ipackqty=request.Form("txt_packqty")(i)
				ijobno=request.Form("txt_jobno")(i)
				if len(trim(ijobno))>0 then
				existPackQty=0
					
					sql="select * from job_pack_detail where box_id='"&boxid&"' and job_number='"&ijobno&"'"
					rs.open sql,conn,1,3
				if not rs.eof then
						existPackQty=csng(rs("packed_qty"))
						
				else
						rs.addnew
						rs("box_id")=boxid
						rs("job_number")=ijobno
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
					rs.update
					rs.close
				
				end if
			next
			
			
			if len(trim(boxId))>0 then
            sql="update job_pack_detail set customer='"&customer&"',remarks='"&remarks&"',customer_label_id='"&custId&"' where box_id ='"&boxId&"'"
			conn.execute(sql)					
		    end if
			
			
			
			
			
			if err.number <> 0 then 
				conn.rollbackTrans   '����ִ�еĲ����ع�
				member("error")="Save failed.����ʧ��.|"&err.description
				response.Write toJSON(member)
			else 
			    conn.commitTrans     'ִ�������ύ
			    member("message")="Save succesful.|����ɹ�."
				response.Write toJSON(member)
			end if	
		case "2"'label print
			opcode=trim(request.Form("txt_opcode"))
			boxid=trim(request.Form("txt_boxid"))
			customer=trim(request.Form("txt_customer"))
			remarks=trim(request.Form("txt_remarks"))
			boxSize=trim(request.Form("txt_boxsize"))			
			packType=trim(request("txt_pack_type"))
			partNo=trim(request.Form("txt_part"))
			packLine=trim(request.Form("txt_pack_line"))
			txt_count=trim(request("txt_count"))
			custId=trim(request("txt_custid"))			
			isBoxid=trim(request("chkBoxidPack"))
			isCustid=trim(request("chkCustidPack"))	
			isPartial=trim(request.Form("chkPartialPack"))
			isTray=trim(request.Form("chkTrayPack"))		
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
			if packType="FG" and qty < cint(boxSize) then
				if isPartial="Y" then
					sql="select user_code from users where user_code='"&opcode&"' and instr(roles_id,'PARTIAL_PACKING')>0 "
					rs.open sql,conn,1,3
					if rs.eof then
						member("error")="You are not authorized to do partial packing.|��û��Ȩ�޽��в��ְ�װ��"
					end if
					rs.close
				else
					member("error")="Packed qty does not equal box size.|��װ����δ�ﵽ����������"
				end if
			end if
				customerPN=""
				customerDef=""
				customerLabel=""
				'����Customer label id
				sql="select customer_pn,customer_define,customer_label,Customer_Config,Customer_pegapn,Customer_desc��nvl(box_size,0) as box_size, nvl(SMALL_PACK,0) as SMALL_PACK,YESNOLITTLELABLE from product_model where item_name='"&partNo&"'"
				rs.open sql,conn,1,3
				if not rs.eof then
					customerPN=rs("customer_pn")
					customerDef=rs("customer_define")
					customerLabel=rs("customer_label")
					Customer_pegapn=rs("Customer_pegapn")
					Customer_desc=rs("Customer_desc")
					Customer_Config=rs("Customer_config")
					boxSize=rs("box_size")
					SMALL_PACK=rs("SMALL_PACK")
					YESNOLITTLELABLE=rs("YESNOLITTLELABLE")
				else
					member("error")="This customer label("&customerLabel&") does not define in system.|�ÿͻ���ǩ("&customerLabel&")��ϵͳ��δ���塣"					
				end if
				rs.close
			 '----------------ֻ��ӡ BoxLable
			 IF isTray<>"Y"  and isCustid<>"Y" Then
					if member("error")="" then
						result=	PrintBoxLable												
						if result <> "" then
							member("error")=result						
						end if	
					end if	 	
			 END IF
			 '----------------------��ӡCustmerLable
			IF isPartial<>"Y" 	and  isTray<>"Y"  then
					if member("error")="" then	
								IF isTray<>"Y" Then
									if member("error")="" then										
											result=PrintCustmerLable
											if result <> "" then
												member("error")=result						
											end if
									end if	
								End If
			        end if
			
			end if
			
			 '------------------ֻ��ӡС��ǩ
			
			IF isPartial<>"Y"  and  isCustid<>"Y"  then			 
							
						 							
									if member("error")="" then
												
												if YESNOLITTLELABLE="YES"  then												
													   if isnull(SMALL_PACK) or  cstr(SMALL_PACK)=0  then
															member("error")="����ϵ����������С��װ��,�����ܴ�С��ǩ"
													   else
															result=	PrintSMALLBoxLable(cstr(boxSize),cstr(SMALL_PACK))										   
													   end if											
												end if
															
												if result <> "" then
													member("error")=result						
												end if			
									end if
								 
						
					
					
			end if		
			
			'------------------��ӡBoxLable �� CustmerLable
			
			IF isPartial="Y"  and isCustid="Y" Then
					if member("error")="" then
					'���ô�ӡ����
								result=	PrintBoxLable												
								if result <> "" then
									member("error")=result						
								end if			
					end if	 	
	
					
					If customerLabel="" then
							member("error")="This part number("&customerPN&") does not define customer label.|���Ϻ�("&customerPN&")δ����ͻ���ǩ��"
					Elseif customerLabel="Tango" then
					
							result=PrintCustmerLable
							if result <> "" then
								member("error")=result						
							end if
					End if
       		END IF	
			
          
	



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
									prodShift="0"
									supplier="0"
								if custId="" then
									countCondition=customerDef&strYear&getOneCharMon(strMon)&strDay&strWeek&N10toC62(cint(right(packLine,1))+9,16)&prodShift&supplier								
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
								
								if member("error")="" then
									
													
									'���ô�ӡ����
									sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"						
									rsAsyn.open sql,conn,1,3						
										if not rsAsyn.eof then
											printName = rsAsyn("PRINTER_NAME")
										end if
									rsAsyn.close
									
									if printName="" then
										alarmMsg="Please contact engiener to set the printer for this machine.|����ϵ����ʦΪ�˻����趨��ȷ�ı�ǩ��ӡ����"
									else
									set PrintCtl=server.createobject("PrintClass.PrintCtl")   
									if   customer<>"Foxconn" then
										
										DateCode=201&strYear&strMon&strDay
										
										rtnCode=PrintCtl.PrintKEBPEGATRONBIGCustomerLabel(printName,Customer_pegapn,customerPN,Customer_Config,Customer_desc,DateCode,boxid,cstr(qty))	
									else             
										rtnCode=PrintCtl.PrintKEBCustomerLabel(printName,custId,cstr(qty),customerPN,partNo,"0.566")
									end if
									'rtnCode="OK"
										if rtnCode<>"OK" then
											alarmMsg="Label print error.��ӡ��ǩ��������.|" & rtnCode		
										end if
								   end if
								end if	
			        PrintCustmerLable=alarmMsg
			   end function
			   
			   function PrintBoxLable			   
			  			 if member("error")="" then
								'���ô�ӡ����
								sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"
								rsAsyn.open sql,conn,1,3
								if not rsAsyn.eof then
									printName = rsAsyn("PRINTER_NAME")
								end if
								rsAsyn.close
								if printName="" then
									alarmMsg="Please contact engiener to set the printer for this machine.|����ϵ����ʦΪ�˻����趨��ȷ�ı�ǩ��ӡ����"
								else
									set PrintCtl=server.createobject("PrintClass.PrintCtl")                  
									if customer<>"Foxconn" then
										rtnCode=PrintCtl.PrintPEGATRONBoxLabel(printName,boxid,partNo,jobno,cstr(qty),remarks,customer)
									else									
										rtnCode=PrintCtl.PrintBoxLabel(printName,boxid,partNo,jobno,cstr(qty),remarks,customer)
									end if
									'rtnCode="OK"
									if rtnCode<>"OK" then
										alarmMsg="Label print error.��ӡ��ǩ��������.|" & rtnCode		
									end if
								end if
						end if	 	
						PrintBoxLable=alarmMsg
			   end function
			   
			   
			   
			   function PrintSMALLBoxLable (txt_count,pouchpack)
					
					Traytime=int(txt_count/pouchpack)
					LastQty=txt_count-pouchpack*Traytime 
					qty=int(pouchpack)
				   '���ô�ӡ����
					sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"						
					rsAsyn.open sql,conn,1,3						
						if not rsAsyn.eof then
						printName = rsAsyn("PRINTER_NAME")
						end if
					rsAsyn.close															
					if printName="" then
						alarmMsg="Please contact engiener to set the printer for this machine.|����ϵ����ʦΪ�˻����趨��ȷ�ı�ǩ��ӡ����"
					else
						set PrintCtl=server.createobject("PrintClass.PrintCtl")                  
						if Traytime>0 then
						for i=1 to Traytime 	
						    if   customer<>"Foxconn" then 
										if  left(custId,7)="RCVJKEB" then
											custIdP= REPLACE(custId,"RCVJKEB","") 
										elseif  left(custId,7)="SPRNKEB" then
											custIdP= REPLACE(custId,"SPRNKEB","") 
										elseif  left(custId,7)="RCVNKEB" then
											custIdP= REPLACE(custId,"RCVNKEB","") 	
										elseif  left(custId,7)="RCVMKEB" then
											custIdP= REPLACE(custId,"RCVMKEB","") 
										else
											custIdP=custId
										end if						 
							rtnCode=PrintCtl.PrintKEBPEGATRONTrayLabel(printName,custIdP,cstr(qty),customerPN,partNo)
							else														
						     rtnCode=PrintCtl.PrintKEBTrayLabel(printName,custId,cstr(qty),customerPN,partNo)
							end if
							'rtnCode="OK"
								if rtnCode<>"OK" then
									alarmMsg="Label print error.��ӡ��ǩ��������.|" & rtnCode		
								end if
						next
						end if
						if LastQty>0 then
						    if customer<>"Foxconn" then 
							rtnCode=PrintCtl.PrintKEBPEGATRONTrayLabel(printName,custIdP,cstr(qty),customerPN,partNo)
							else
							rtnCode=PrintCtl.PrintKEBTrayLabel(printName,custId,cstr(LastQty),customerPN,partNo)
							end if
							'rtnCode="OK"
								if rtnCode<>"OK" then
									alarmMsg="Label print error.��ӡ��ǩ��������.|" & rtnCode		
								end if															
						end if
					end if
					PrintSMALLBoxLable=alarmMsg
			   end function

			if err.number <> 0 then 
						conn.rollbackTrans   '����ִ�еĲ����ع�
						member("error")=err.description
					else 
						conn.commitTrans     'ִ�������ύ
			end if	
				
			member("cust_id")=custId	
				
	
			response.Write toJSON(member)
		case "3" 'check 2D code
			opcode=trim(request.Form("txt_opcode"))
			codeValue=trim(request("txt_2DCode"))
			boxId=trim(request("txt_boxid"))
			packType=trim(request.Form("txt_pack_type"))
			partNo=trim(request.Form("txt_part"))
			packLine=trim(request.Form("txt_pack_line"))
			boxSize=trim(request.Form("txt_boxsize"))
			prodShift=trim(request.Form("txt_shift"))
			supplier=trim(request.Form("txt_supplier"))
			codeLen=len(codeValue)
			FACTORY_CODE=left(codeValue,3)
			codeDate=mid(codeValue,4,4)
			codeLine=mid(codeValue,8,1)
			CODE_NAME=mid(codeValue,12,4) 
            VERSION_NUMBER=mid(codeValue,16,1)			
			startime=timer()	
			conn.beginTrans
			'sql="select a.job_number,a.defect_code_id,a.box_id,a.is_packed,b.part_number_tag,b.line_name,b.shift,b.supplier from job_2d_code a,job b "
			'sql=sql+" where a.job_number=b.job_number and a.sheet_number=b.sheet_number and a.code='"&codeValue&"'" 
			sql="select  get_sub_job_number(a.job_number, a.sheet_number) as subJob, a.job_number,a.sheet_number,a.defect_code_id,a.box_id,a.is_packed,b.part_number_tag,b.line_name,b.shift,b.supplier from job_2d_code a,job b "
			sql=sql+" where a.job_number=b.job_number and  a.code='"&codeValue&"'" 
			rsAsyn.open sql,conn,1,3
			'member("error")=sql
			'response.End()
			if rsAsyn.eof then
				member("error")="This 2D Code("&codeValue&") dose not exist.|ɨ��Ķ�ά��("&codeValue&")������."
			else
	
					if member("error")="" and packType="FG"  then
						job_number=rsAsyn("job_number")
						sql="select job_number from job_master_store_pre where job_number='"&rsAsyn("job_number")&"' and instr(sub_job_numbers,'"&rsAsyn("subJob")&"')>0 "
						rs.open sql,conn,1,3
						if rs.eof then
						member("error")="This 2D Code("&codeValue&") has not done pre store .|ɨ��Ķ�ά��("&codeValue&")δ��Ԥ���."
						end if
						rs.close
					end if	
			
			   sql="select * from  job where job_number='"&rsAsyn("job_number")&"' and  sheet_number='"&rsAsyn("sheet_number")&"'"
			   rs.open sql,conn,1,3
			   if rs.eof then
			   member("error")="�˹���("&rsAsyn("job_number")&")û�п�վ����ֱ�Ӵ��."
			   end if
			   rs.close
				defectId=rsAsyn("DEFECT_CODE_ID")	
				if rsAsyn("is_packed") = "1" then
					member("error")="This 2D Code("&codeValue&") was packed.|ɨ��Ķ�ά��("&codeValue&")�Ѿ���װ��."
				elseif defectId <> "" and packType="FG" then
					member("error")="This 2D Code("&codeValue&") was scraped, cannot do FG packing.|ɨ��Ķ�ά��("&codeValue&")�Ѿ����ϣ���������Ʒ��װ."
					
				elseif isnull(defectId) and packType="SCRAP" then
					member("error")="This 2D Code("&codeValue&") is good, cannot do Scrap packing.|ɨ��Ķ�ά��("&codeValue&")����Ʒ�����������ϰ�װ."	
			'	elseif left(rsAsyn("job_number"),2)<>"KB"  and left(rsAsyn("job_number"),3)<>"EKB" then
		
					'	member("error")="This Job("&rsAsyn("job_number")&") is not KB, cannot do FG packing.|ɨ��Ķ�ά��("&codeValue&")�Ĺ�������������KB)��������������Ʒ��װ."		
				
				end if
				
				if member("error")="" and packType="FG" or PackType="SCRAP" then
					if partNo = "" then
						partNo=rsAsyn("part_number_tag")
					elseif partNo <> rsAsyn("part_number_tag") then
						member("error")="This 2D Code("&codeValue&") is mix packing part number.|�ö�ά����Ϻ�("&rsAsyn("part_number_tag")&")���Ѱ�װ���ϺŲ�һ��."					
					end if
					
					if packLine = "" then
						packLine=rsAsyn("line_name")
					elseif packLine<>rsAsyn("line_name") then
						member("error")="This 2D Code("&codeValue&") is mix packing line.|�ö�ά����߱�("&rsAsyn("line_name")&")���Ѱ�װ���߱�һ��."
					end if
			
				'	if prodShift = "" then
				'		prodShift=rsAsyn("shift")
				'	elseif prodShift<>rsAsyn("shift") then
				'		member("error")="This 2D Code("&codeValue&") is mix product shift.|�ö�ά����������("&rsAsyn("shift")&")���Ѱ�װ�Ĳ�һ��."
				'	end if
					if supplier = "" then
						supplier=rsAsyn("supplier")
					elseif supplier<>rsAsyn("supplier") then
						member("error")="This 2D Code("&codeValue&") is mix supplier.|�ö�ά��Ĺ�Ӧ��("&rsAsyn("supplier")&")���Ѱ�װ�Ĳ�һ��."
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
					'���Pack Qty�Ƿ񳬹�Job ��Good Qty
					
					if member("error")="" and packType="FG"  then
					
					
	                SQL="select * from Line where Line_name='"&packLine&"'"
                    rs.open SQL,conn,1,3
                    if not rs.eof then
					FACTORY_CODEO=rs("FACTORY_CODE")
					CODE_LINENAMEO=rs("CODE_LINENAME")
					CODE_NAMEO=rs("CODE_NAME")
					VERSION_NUMBERO=rs("VERSION_NUMBER")
					CODE_DateO=rs("CODE_Date")
					end if
					rs.close
					
				' member("error")="��Ʒ�벻��("&CODE_LINENAMEO&")���˲�Ʒ���ܳ�����"
				'if 	instr(FACTORY_CODE,FACTORY_CODEO)="0" then
				' member("error")="�Բ��𣡱�����������ӦΪ:DYD"
				'end if
				
				  ' ��ߵ��ַ��Ƿ���ǰ�ߵ��ַ���
				
				
				if	instr(CODE_LINENAMEO,codeLine)=0 then
				member("error")="�Բ����߱𲻶ԣ�����Ʒ�߱���("&codeLine&")��Ӧ����("&CODE_LINENAMEO&"��"
				end if
					
				
				
				if	instr(VERSION_NUMBERO,VERSION_NUMBER)=0 then
				member("error")="�汾�Ų��ԣ��˲�Ʒ�汾("&VERSION_NUMBER&")��Ӧ����("&VERSION_NUMBERO&")"
				end if
				
				if 	instr(CODE_NAMEO,CODE_NAME)=0 then
				member("error")="��Ʒ�벻��("&CODE_NAME&")���˲�Ʒ���ܳ�����Ӧ����("&CODE_NAMEO&"��"
				end if	
				
				
				if CODE_DateO<>""  then
					if codeDate<=CODE_DateO then
					member("error")="��Ʒ��������("&codeDate&")���˲�Ʒ���ܳ�����"
					end if
				end if
				
				
					
					goodQty=0
					sql="select good_qty from job_inventory where job_number in ('"&rsAsyn("job_number")&"')"
					rs.open sql,conn,1,3
					if not rs.eof then
						goodQty =CLng( rs("GOOD_QTY"))
					end if
					rs.close
								
					'check pack qty

					qty=0			
					for i=1 to request.Form("txt_packqty").count
						if len(trim(request.Form("txt_packqty")(i)))>0 then
							ipackqty=trim(request.Form("txt_packqty")(i))
							ijobno=trim(request.Form("txt_jobno")(i))
							'if ijobno=rsAsyn("job_number") and cint(ipackqty)>(goodQty-1) then
							    if goodQty = 0 then
								member("error")="The pack qty("&ipackqty&") of "&ijobno&" exceeds good qty("&goodQty&")|"&ijobno&"�İ�װ����("&ipackqty&")������Ʒ����("&goodQty&")"
								response.Write toJSON(member)
								response.End()
							    end if
							qty=qty+cint(request.Form("txt_packqty")(i))
						end if
					next
					
					end if
					
					    if (qty+1) >boxSize then
						member("error")="Total pack qty("&qty&") exceeds box qty("&boxSize&").|��װ����("&qty&")����ֽ���װ������("&boxSize&")."
					    end if
				        set rs1=server.CreateObject("adodb.recordset")
						sqlc="select COUNT(*) P_Count from job_2d_code where box_id='"&boxId&"'" 
						rs1.open sqlc,conn,1,3
						if  rs1("P_Count") <> "0" then
						P_Count=rs1("P_Count")
				        end if
				       if cint(P_Count) >= cint(boxSize) then
				       member("error")="��װ��������ֽ���װ������("&boxSize&")."
			           end if
				
		      end if
				
				
				
				
				if member("error")="" then	
				    if  packType="FG" then		
					result=checkTestResult(codeValue,rsAsyn("part_number_tag"))	
					end if				
					if result <> "" then
						member("error")=result						
					else
						member("job_number")=rsAsyn("job_number")
						member("part_number")=rsAsyn("part_number_tag")
						member("pack_line")=packLine
						member("boxsize")=boxSize
						member("prod_shift")=prodShift
						member("txt_supplier")=supplier	
						if boxId="" then
							'����box id						
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
						PACK_TIME=now()
						sql="update job_2d_code set box_id='"&boxId&"',is_packed=1,PACK_TIME='"&PACK_TIME&"',PACK_USER='"&opcode&"' where code='"&codeValue&"'" 
						conn.execute(sql)
					    set rs1=server.CreateObject("adodb.recordset")
						sqlc="select COUNT(*) P_Count from job_2d_code where box_id='"&boxId&"'" 
						rs1.open sqlc,conn,1,3
						if  rs1("P_Count") <> "0" then
						
						P_Count=rs1("P_Count")
						
						member("PCount")=P_Count
						else
					    member("PCount")="0"
						end if
						rs1.close
						JS=1
					sql="update JOB_INVENTORY set PACKED_QTY=decode(PACKED_QTY,null,0,PACKED_QTY)+"&JS 
					if packType="FG" then
						sql=sql+",GOOD_QTY=GOOD_QTY-"&JS
					else
						sql=sql+",SCRAP_QTY=SCRAP_QTY-"&JS	
					end if
					sql=sql+" where JOB_NUMBER ='"&job_number&"'"
					
					conn.execute(sql)					
						
						
						
					end if
				end if
			end if
			rsAsyn.close
			if err.number <> 0 then 
				conn.rollbackTrans   '����ִ�еĲ����ع�
				member("error")=err.description
			else 
			    conn.commitTrans     'ִ�������ύ			    				
			end if
			endtime=timer()				
		    member("OKT")=cstr(FormatNumber((endtime-startime),2,-1))
			response.Write toJSON(member)
	end select
	set rsAsyn=nothing 	
end if

function checkTestResult(str2DCode,strPartNumber)
	testName=""
	'get unit test result
	'strSql="select linename,adfail from ( "
	'strSql=strSql+"select a.ad_id, linename, a.measuredatetime, adfail,ROW_NUMBER() over(partition by a.linename order by a.measuredatetime desc) rn "
	'strSql=strSql+" from pvs.vw_adid_by_sn a,pvs.ad_serial b where a.ad_id=b.ad_id and b.serialnumber='"&str2DCode&"' "
	'strSql=strSql+") temp  where  rn <10"
	'strSql="select testResult,TEST_ENVIRONMENT from  TEST_RESULT_SUMMARY a left join TEST_RESULT_MAIN_KEB b on(a.archive_id=b.ARCHIVE_ID) where serialnumber = '"&str2DCode&"'" 
 'strSql=" select testResult,TEST_ENVIRONMENT from  TEST_RESULT_SUMMARY a,TEST_RESULT_MAIN_KEB b  where a.archive_id=b.ARCHIVE_ID and serialnumber = '"&str2DCode&"'" 
	 strSql="select  a.testResult,a.testitem from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&str2DCode&"'" 
	rsTSD.open strSql,connTSD,1,3	
	K=0
	if rsTSD.eof then
	alarmMsg="��ѧ����δȫ��ͨ������ע��鿴����"	
	else
	while not rsTSD.eof 	
	   
	
		if rsTSD("testResult")="PASS" then
			testName=testName&rsTSD("testitem")&","

		else
			
		alarmMsg="��ѧ����δȫ��ͨ������ע��鿴����"	
		end if
		rsTSD.movenext
	wend
	end if
	rsTSD.close
	if testName<>"" then
		testName=left(testName,len(testName)-1)
	end if
	
	
	if alarmMsg="" then
	alarmMsg=""
	set rsPart=server.CreateObject("adodb.recordset")
	strsqlPart="select test_name,alarm_msg from part_test_name where part_number = '"&strPartNumber&"' and test_type='TSDFG'"
	'alarmMsg=strsqlPart
	rsPart.open strsqlPart,conn,1,3 
	if not rsPart.eof then
	
	
	
	strSql="select test_name,alarm_msg from part_test_name where part_number = '"&strPartNumber&"' and test_type='TSDFG' and test_name not in('"&replace(testName,",","','")&"') "
	rs.open strSql,conn,1,3
	
	while not rs.eof
		'alarmMsg=alarmMsg&rs("alarm_msg")&"|"
		alarmMsg=alarmMsg&rs("test_name")&"��"&rs("alarm_msg")&"|"
		rs.movenext
	wend
	rs.close
	else
	
	alarmMsg="("&strPartNumber&")���Ϻ�δ���а�װ��������ϵ����."
	
	end if
	rsPart.close
	'if k<2 then
	'alarmMsg="һ����ѧ����С��2�Σ�"
	'end if
	if alarmMsg<>"" then
		alarmMsg=left(alarmMsg,len(alarmMsg)-1)
	end if
	end if
	checkTestResult=alarmMsg
end function
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/TSD_Close.asp" -->