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
			opcode=request.Form("txt_opcode")
			boxid=request.Form("txt_boxid")
			customer=request.Form("txt_customer")
			remarks=request.Form("txt_remarks")
			packType=request("txt_pack_type")
			packLine=trim(request.Form("txt_pack_line"))
			partnumber=trim(request.Form("txt_part"))
			prodShift=trim(request.Form("txt_shift"))
			supplier=trim(request.Form("txt_supplier"))
			
			conn.beginTrans
			for i=1 to request.Form("txt_jobno").count
				'更新Job的Good Qty=Good Qty - Pack Qty，Packed Qty = Packed Qty + Pack Qty
				ipackqty=request.Form("txt_packqty")(i)
				ijobno=request.Form("txt_jobno")(i)
				if len(trim(ijobno))>0 then
					sql="update JOB_INVENTORY set PACKED_QTY=decode(PACKED_QTY,null,0,PACKED_QTY)+"&ipackqty 
					if packType="FG" then
						sql=sql+",GOOD_QTY=GOOD_QTY-"&ipackqty
					else
						sql=sql+",SCRAP_QTY=SCRAP_QTY-"&ipackqty	
					end if
					sql=sql+" where JOB_NUMBER ='"&ijobno&"'"
					conn.execute(sql)
					'新增包装的记录到JOB_PACK_DETAIL
					sql="insert into JOB_PACK_DETAIL (BOX_ID,JOB_NUMBER,PART_NUMBER,PACKED_QTY,CUSTOMER,PACKED_USER,PACKED_TIME,REMARKS,PACKED_LINE,PROD_SHIFT,SUPPLIER) values ('"&boxid&"','"&ijobno&"','"&partnumber&"','"&ipackqty&"','"&customer&"','"&opcode&"',sysdate,'"&remarks&"','"&packLine&"','"&prodShift&"','"&supplier&"')"
					conn.execute(sql)
				end if
			next
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
			opcode=request.Form("txt_opcode")
			boxid=request.Form("txt_boxid")
			customer=request.Form("txt_customer")
			remarks=request.Form("txt_remarks")
			boxSize=request.Form("txt_boxsize")
			isPartial=request.Form("chkPartialPack")
			packType=request("txt_pack_type")
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
						member("error")="You are not authorized to do partial packing.|您没有权限进行部分包装。"
					end if
					rs.close
				else
					member("error")="Packed qty does not equal box size.|包装数量未达到满箱数量。"
				end if
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
					rtnCode=PrintCtl.PrintBoxLabel(printName,boxid,partnumber,jobno,cstr(qty),remarks,customer)
					'rtnCode="OK"
					if rtnCode<>"OK" then
						member("error")="Label print error.打印标签发生错误.|" & rtnCode		
					end if
				end if
			end if			
			response.Write toJSON(member)
		case "3" 'check 2D code
			codeValue=trim(request("txt_2DCode"))
			boxId=trim(request("txt_boxid"))
			packType=trim(request.Form("txt_pack_type"))
			partNo=trim(request.Form("txt_part"))
			packLine=trim(request.Form("txt_pack_line"))
			boxSize=trim(request.Form("txt_boxsize"))
			prodShift=trim(request.Form("txt_shift"))
			supplier=trim(request.Form("txt_supplier"))
			
			conn.beginTrans
			sql="select a.job_number,a.defect_code_id,a.box_id,a.is_packed,b.part_number_tag,b.line_name,b.shift,b.supplier from job_2d_code a,job b "
			sql=sql+" where a.job_number=b.job_number and a.sheet_number=b.sheet_number and a.code='"&codeValue&"'" 
			rsAsyn.open sql,conn,1,3
			if rsAsyn.eof then
				member("error")="This 2D Code("&codeValue&") dose not exist.|扫描的二维码("&codeValue&")不存在."
			else
				defectId=rsAsyn("DEFECT_CODE_ID")	
					
				if rsAsyn("is_packed") = "1" then
					member("error")="This 2D Code("&codeValue&") was packed.|扫描的二维码("&codeValue&")已经包装过."
				elseif defectId <> "" and packType="FG" then
					member("error")="This 2D Code("&codeValue&") was scraped, cannot do FG packing.|扫描的二维码("&codeValue&")已经报废，不能做良品包装."
				elseif isnull(defectId) and packType="SCRAP" then
					member("error")="This 2D Code("&codeValue&") is good, cannot do Scrap packing.|扫描的二维码("&codeValue&")是良品，不能做报废包装."				
				end if
				
				if member("error")="" and packType="FG" then
					if partNo = "" then
						partNo=rsAsyn("part_number_tag")
					elseif partNo <> rsAsyn("part_number_tag") then
						member("error")="This 2D Code("&codeValue&") is mix packing part number.|该二维码的料号("&rsAsyn("part_number_tag")&")与已包装的料号不一致."					
					end if
					
					if packLine = "" then
						packLine=rsAsyn("line_name")
					elseif packLine<>rsAsyn("line_name") then
						member("error")="This 2D Code("&codeValue&") is mix packing line.|该二维码的线别("&rsAsyn("line_name")&")与已包装的线别不一致."
					end if
					
				'	if prodShift = "" then
				'		prodShift=rsAsyn("shift")
				'	elseif prodShift<>rsAsyn("shift") then
				'		member("error")="This 2D Code("&codeValue&") is mix product shift.|该二维码的生产班别("&rsAsyn("shift")&")与已包装的不一致."
				'	end if
					if supplier = "" then
						supplier=rsAsyn("supplier")
					elseif supplier<>rsAsyn("supplier") then
						member("error")="This 2D Code("&codeValue&") is mix supplier.|该二维码的供应商("&rsAsyn("supplier")&")与已包装的不一致."
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
					'检查Pack Qty是否超过Job 的Good Qty
					goodQty=0
					sql="select good_qty from job_inventory where job_number in ('"&rsAsyn("job_number")&"')"
					rs.open sql,conn,1,3
					if not rs.eof then
						goodQty = cint(rs("GOOD_QTY"))
					end if
					rs.close
										
					'check pack qty
					qty=0			
					for i=1 to request.Form("txt_packqty").count
						if len(trim(request.Form("txt_packqty")(i)))>0 then
							ipackqty=trim(request.Form("txt_packqty")(i))
							ijobno=trim(request.Form("txt_jobno")(i))
							if ijobno=rsAsyn("job_number") and cint(ipackqty)>(goodQty-1) then
								member("error")="The pack qty("&ipackqty&") of "&ijobno&" exceeds good qty("&goodQty&")|"&ijobno&"的包装数量("&ipackqty&")超过良品数量("&goodQty&")"
								response.Write toJSON(member)
								response.End()
							end if
							qty=qty+cint(request.Form("txt_packqty")(i))
						end if
					next
					if (qty+1) >boxSize then
						member("error")="Total pack qty("&qty&") exceeds box qty("&boxSize&").|包装数量("&qty&")超过纸箱可装的数量("&boxSize&")."
					end if
					
				end if
				
				if member("error")="" then					
					result=checkTestResult(codeValue,rsAsyn("part_number_tag"))					
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
							'生成box id						
							countType="BoxId"
							countCondition=packType&packLine&formatdate(Now,"ymmdd")
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
								boxId=countCondition&repeatstring(rs("count_value"),"0",3)
								rs("count_value")=clng(rs("count_value"))+1
								rs("lm_time")=now()
							end if
							rs.update
							rs.close													
						end if
						member("boxid")=boxId
						sql="update job_2d_code set box_id='"&boxId&"',is_packed=1 where code='"&codeValue&"'" 
						conn.execute(sql)
					end if
				end if
			end if
			rsAsyn.close
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

function checkTestResult(str2DCode,strPartNumber)
	testName=""
	'get unit test result
	strSql="select linename,adfail from ( "
	strSql=strSql+"select a.ad_id, linename, a.measuredatetime, adfail,ROW_NUMBER() over(partition by a.linename order by a.measuredatetime desc) rn "
	strSql=strSql+" from pvs.vw_adid_by_sn a,pvs.ad_serial b where a.ad_id=b.ad_id and b.serialnumber='"&str2DCode&"' "
	strSql=strSql+") temp  where  rn <2 "
	
	rsPVS.open strSql,connPVS,1,3	
	K=0
	while not rsPVS.eof 	
		if rsPVS("adfail")="False" then
			testName=testName&left(rsPVS("linename"),3)&","
		 
		  if  left(rsPVS("linename"),3)="AF0" then
		     k=K+1
		  end if 	
		else
			
		alarmMsg="声学测试未全部通过，请注意查看！！"	
		end if
		rsPVS.movenext
	wend
	rsPVS.close
	if testName<>"" then
		testName=left(testName,len(testName)-1)
	end if
	
	
	if alarmMsg="" then
	alarmMsg=""
	strSql="select test_name,alarm_msg from part_test_name where part_number = '"&strPartNumber&"' and test_name not in('"&replace(testName,",","','")&"') "
	rs.open strSql,conn,1,3
	while not rs.eof
		alarmMsg=alarmMsg&rs("alarm_msg")&"|"
		rs.movenext
	wend
	rs.close
	if alarmMsg<>"" then
		alarmMsg=left(alarmMsg,len(alarmMsg)-1)
	end if
	end if
	if alarmMsg="" then
	if k>2 then
	alarmMsg="一次声学测试小于2次！"
	end if
	end if
	checkTestResult=alarmMsg
end function
%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/PVS_Close.asp" -->