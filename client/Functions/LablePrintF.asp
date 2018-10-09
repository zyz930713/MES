
<%
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
									countCondition=VENDOR_CODE&strYear&getOneCharMon(strMon)&strDay&strWeek&N10toC62(cint(right(packLine,1))+9,16)&prodShift&supplier								
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
										alarmMsg="Please contact engiener to set the printer for this machine.打印标签发生错误，请联系主管"
									else
									set PrintCtl=server.createobject("PrintClass.PrintCtl")   
									if   customer<>"Foxconn" then
										
										DateCode=201&strYear&strMon&strDay
										
										rtnCode=PrintCtl.PrintKEBPEGATRONBIGCustomerLabel(printName,OEM_PN,customerPN,REV,Customer_Config,CUSTOMER_DESC_EN,DateCode,boxid,cstr(txt_count),CUSTOMER_DESC_CHINA)	
									else             
										rtnCode=PrintCtl.PrintKEBCustomerLabel(printName,custId,cstr(qty),customerPN,partNo,CUSTOMER_DESC_EN,"0.566")
									end if
									'rtnCode="OK"
										if rtnCode<>"OK" then
											alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
										end if
								   end if
								end if	
			        PrintCustmerLable=alarmMsg
 end function
			   
function PrintBoxLable			   
			  			 if member("error")="" then
								'调用打印功能
								sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"
								rsAsyn.open sql,conn,1,3
								if not rsAsyn.eof then
									printName = rsAsyn("PRINTER_NAME")
								end if
								rsAsyn.close
								if printName="" then
									alarmMsg="Please contact engiener to set the printer for this machine.打印标签发生错误，请联系主管"
								else
									set PrintCtl=server.createobject("PrintClass.PrintCtl")                  
									if customer<>"Foxconn" then
										rtnCode=PrintCtl.PrintPEGATRONBoxLabel(printName,boxid,partNo,jobno,cstr(txt_count),remarks,customer)
									else									
										rtnCode=PrintCtl.PrintBoxLabel(printName,boxid,partNo,jobno,cstr(txt_count),remarks,customer)
									end if
									'rtnCode="OK"
									if rtnCode<>"OK" then
										alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
									end if
								end if
						end if	 	
						PrintBoxLable=alarmMsg
end function
			   
			   
			   
function PrintSMALLBoxLable (txt_count,pouchpack)
					
					Traytime=int(txt_count/pouchpack)
					LastQty=txt_count-pouchpack*Traytime 
					qty=int(pouchpack)
				   '调用打印功能
					sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request.Form("computername"))+"'"						
					rsAsyn.open sql,conn,1,3						
						if not rsAsyn.eof then
						printName = rsAsyn("PRINTER_NAME")
						end if
					rsAsyn.close															
					if printName="" then
						alarmMsg="Please contact engiener to set the printer for this machine.打印标签发生错误，请联系主管"
					else
						set PrintCtl=server.createobject("PrintClass.PrintCtl")                  
						if Traytime>0 then
						for i=1 to Traytime 	
						    if   customer<>"Foxconn" then 
										
										custIdP=right(custId,13)			 
							rtnCode=PrintCtl.PrintKEBPEGATRONTrayLabel(printName,custIdP,cstr(qty),customerPN,partNo)
							else														
						     rtnCode=PrintCtl.PrintKEBTrayLabel(printName,custId,cstr(qty),customerPN,CUSTOMER_DESC_EN,partNo)
							end if
							'rtnCode="OK"
								if rtnCode<>"OK" then
									alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
								end if
						next
						end if
						if LastQty>0 then
						    if customer<>"Foxconn" then 
							custIdP=right(custId,13)
							rtnCode=PrintCtl.PrintKEBPEGATRONTrayLabel(printName,custIdP,cstr(LastQty),customerPN,partNo)
							else
							rtnCode=PrintCtl.PrintKEBTrayLabel(printName,custId,cstr(LastQty),customerPN,CUSTOMER_DESC_EN,partNo)
							end if
							'rtnCode="OK"
								if rtnCode<>"OK" then
									alarmMsg="Label print error.打印标签发生错误.|" & rtnCode		
								end if															
						end if
					end if
					PrintSMALLBoxLable=alarmMsg
 end function
%>