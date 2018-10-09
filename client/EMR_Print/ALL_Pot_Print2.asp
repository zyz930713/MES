<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Ticket_Open.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>KEB Barcode System</title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
</head>
<script language="javascript"  type="text/javascript">
	 function Print()
	 {
		 
		 if(document.form1.JOB_NUMBER.value=="")
		{
			window.alert("请输入工单号！");
			document.form1.JOB_NUMBER.focus();
			return;
		}
		  
		   if(document.form1.Inspector.value=="")
		{
			window.alert("请输入检查人！");
			document.form1.Inspector.focus();
			return;
		}
		  
		     if(document.form1.BATCHNOA.value=="")
		{
			window.alert("Batch no 不能为空！");
			document.form1.BATCHNOA.focus();
			return;
		} 
		
		    if(document.form1.BATCHNOB.value=="")
		{
			window.alert("Batch no 不能为空！");
			document.form1.BATCHNOB.focus();
			return;
		} 
		
		  if(document.form1.Qty.value>document.form1.PACKED_QTY.value)
		  {
			  window.alert("数量不能大于默认数值");
			document.form1.Qty.focus();
			return;
		}
		 if (!document.form1.Equipment_NO.value)
   		 {
    		alert("请输入机器号！");
			document.form1.Equipment_NO.focus();
			return false;	
		
		}		
		else{
			
			form1.action="ALL_Pot_print2.asp?Action=2";
			form1.submit();
		}
	 }
</script>
<%
		computername=trim(request("computername"))
		part_NUMBER=trim(request("PART_NUMBER"))
		packed_USER=trim(request("PACKED_USER"))
		 set rs1=server.createobject("adodb.recordset")
		 sql="select * from Product_model where item_name='"&PART_NUMBER&"'"
			rs1.open SQL,conn,1,3
			if not rs1.eof then	
				PRODUCTION_NAME=rs1("Description")
				PACKED_QTY=rs1("Box_size")
				SMALL_pack=rs1("SMALL_pack")
				CCL=rs1("CCL")
				if isnull(CCL) then
					word="请配置线别（CCL）。"
					response.Redirect("Pot_print.asp?word="&word)
				end if
				if isnull(SMALL_pack) then
					word="请配置最小包装数量。"
					response.Redirect("Pot_print.asp?word="&word)
				end if
			else
				word="料号不存在，请输入正确的料号"		
				response.Redirect("Pot_print.asp?word="&word)
			end if
			rs1.close
			set rs4=nothing
			Action=request("Action")
			If Action="2" then
			    conn.begintrans
			    set rsAsyn=server.CreateObject("adodb.recordset")
					PRODUCTION_NAME=trim(request("PRODUCTION_NAME"))				
					boxId=left(trim(request("Box_id")),12)
					Shift=trim(request("Shift"))
					pack_Time=trim(request("pack_Time"))
					qty=trim(request("qty"))
					Top_Plate=trim(request("Top_Plate"))
					Magent=trim(request("Magent"))				
					Equipment_NO=trim(request("Equipment_NO"))
					Inspector=trim(request("Inspector"))
					Remark=trim(request("Remark"))
					BATCHNOA=trim(request("BATCHNOA"))
					BATCHNOB=trim(request("BATCHNOB"))
					JOB_NUMBER=trim(request("JOB_NUMBER"))
					
					
				SQLH="SELECT * FROM tbl_MES_LotMaster  where WipEntityName='"&JOB_NUMBER&"'"
				rsTicket.open SQLH,connTicket,1,3

					if  rsTicket.bof and rsTicket.eof then
					  '  word="此工单不存在，请确认工单。"						
						'response.Redirect("Pot_print.asp?word="&word)
					END IF	
				rsTicket.close
				
				
				sqlA="select jsq from EMR_PACK_DETAIL where Box_id='"&left(BATCHNOA,12)&"'"
				
				rs.open sqlA,conn,1,3
				
				if not (rs.bof and rs.eof ) then
				
				    jsqA=rs("jsq")
					
				    if  isnull(jsqA) then
						word="此箱号里不是外磁箱号。"						
						response.Redirect("Pot_print.asp?word="&word)	
					elseif  int(jsqA)<=0 then
						word=BATCHNOA&"此外磁箱号已使用过，不能重复使用。"						
						response.Redirect("Pot_print.asp?word="&word)
					end if
				else
				 	word="此箱号不存在。"						
					response.Redirect("Pot_print.asp?word="&word)
				end if
				rs.close
				
				sqlA="select jsq from EMR_PACK_DETAIL where Box_id='"&left(BATCHNOB,12)&"'"
				
				rs.open sqlA,conn,1,3
				
				if not (rs.bof and rs.eof ) then
					jsqB=rs("jsq")
					if  isnull(jsqB) then
						word="此箱号里不是内磁箱号。"						
						response.Redirect("Pot_print.asp?word="&word)	
					elseif  int(jsqB)<=0 then
						word=BATCHNOB&"此内磁箱号已使用过，不能重复使用。"						
						response.Redirect("Pot_print.asp?word="&word)
					end if
				else
				 	word="此箱号不存在。"						
					response.Redirect("Pot_print.asp?word="&word)	
				end if
				rs.close
								
					if boxId="" then
							'生成box id						
							countType="POT"
							countCondition=CCL&formatdate(Now,"ymmdd")
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
				'调用打印功能
				sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request("computername"))+"'"
				rsAsyn.open sql,conn,1,3
				if not rsAsyn.eof then
					printName = rsAsyn("PRINTER_NAME")
				end if
				rsAsyn.close
				if printName="" then
					word="Please contact engiener to set the printer for this machine.|请联系工程师为此机器设定正确的标签打印机。"
					
					response.Redirect("Pot_print.asp?word="&word)
				else
					set PrintCtl=server.createobject("PrintClass.PrintCtl")
					
					JSQ=int(cint(qty)/cint(SMALL_pack))
				    LastQty=cint(cint(qty)-cint(SMALL_pack)*JSQ)
					if JSQ>0 then
						for i=1 to JSQ 
						pack_TimeL=formatdate(now,"yyyymmdd")
						boxIdN=boxId&"-"&i 
					    rtnCode=PrintCtl.PrintKEBALLPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxIdN,Shift,pack_TimeL,SMALL_pack,BATCHNOA,BATCHNOB,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark)
						TimeDelaySeconds(1)
						'rtnCode="OK"
						if rtnCode<>"OK" then
							word="Label print error.打印标签发生错误.|" & rtnCode		
							response.Redirect("Pot_print.asp?word="&word)
						end if
						next
					end if	
				if LastQty>0 then
				boxIdN=boxId&"-"&i
					rtnCode=PrintCtl.PrintKEBALLPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxIdN,Shift,pack_TimeL,LastQty,BATCHNOA,BATCHNOB,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark)
					'rtnCode="OK"
					if rtnCode<>"OK" then
				            word="Label print error.打印标签发生错误.|" & rtnCode		
							response.Redirect("Pot_print.asp?word="&word)		
				    end if
				end if
				   	rtnCode=PrintCtl.PrintKEBALLPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxId,Shift,pack_Time,qty,BATCHNOA,BATCHNOB,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark)
					'rtnCode="OK"
				    if rtnCode="OK" then
					PACK_TIME=now()
					sql="insert into EMR_PACK_DETAIL  (PACK_USER,PART_NUMBER,PRODUCTION_NAME,BOX_ID,SHIFT,PACK_TIME,PACK_QTY,EQUIPMENT_NO,INSPECTOR,REMARK,JOB_NUMBER,BATCHNOA,BATCHNOB) VALUES  ('"+packed_USER+"','"+part_NUMBER+"','"+PRODUCTION_NAME+"','"+boxId+"','"+Shift+"','"&PACK_TIME&"','"+qty+"','"+EQUIPMENT_NO+"','"+INSPECTOR+"','"+REMARK+"','"+JOB_NUMBER+"','"+BATCHNOA+"','"+BATCHNOB+"')"
					conn.execute(sql)
					
					JSQA=JSQA-qty
					JSQB=JSQB-qty
					
					sql="update EMR_PACK_DETAIL set jsq='"&JSQA&"' where box_id='"&left(BATCHNOA,12)&"' "
					conn.execute(sql)
					sql="update EMR_PACK_DETAIL set jsq='"&JSQB&"' where box_id='"&left(BATCHNOB,12)&"' "
					conn.execute(sql)
					else
						word="Label print error.打印标签发生错误.|" & rtnCode		
					end if
				end if
				if err.number>0 then
					conn.rollbacktrans
					errMsg="打印失败|"&err.description
					response.Redirect("Pot_print.asp?word="&errMsg)	
				else 
					conn.commitTrans     '执行事务提交
				end if	
			end if
Sub TimeDelaySeconds(DelaySeconds) 
SecCount = 0 
Sec2 = 0 
While SecCount < DelaySeconds + 1 
Sec1 = Second(Time()) 
If Sec1 <> Sec2 Then 
Sec2 = Second(Time()) 
SecCount = SecCount + 1 
End If 
Wend 
End Sub 
%>

<body onLoad="form1.txtSubJobList.focus();" bgcolor="#339966">

<table width="1200" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">  	
    <tr>
      <td height="20" colspan="6" class="t-t-DarkBlue"  align="center">(Hektor&amp;Maple Pot Assy Print)</td>
    </tr>
	
    <tr>
      <td width="152" rowspan="2" >Operator Code </td>
      <td width="328" rowspan="2" ><input name="PACKED_USER" type="text" id="PACKED_USER" value=<%=PACKED_USER%>  readonly="true" style="background-color:#666666"></td>
      <td width="150" >Inspector</td>
      <td colspan="3" ><label for="select">
        <input name="Inspector" type="text" id="Inspector" value=<%=Inspector%> >
      </label></td>
    </tr>
    <tr>
      <td height="25" >Description</td>
      <td colspan="2" >Batch No</td>
      <td width="197" >&nbsp;</td>
    </tr>
    <tr>
      <td >Production</td>
      <td ><input name="PRODUCTION_NAME" type="text" id="PRODUCTION_NAME" value="<%=PRODUCTION_NAME%>" size="40" readonly style="background-color:#666666" ></td>
      <td >Outer Pot</td>
      <td colspan="2" ><input name="BATCHNOA" type="text" id="BATCHNOA" value="<%=BATCHNOA%>" size="30" ></td>
      <td >&nbsp;</td>
    </tr>
    <tr>
      <td >12NC</td>
      <td ><input name="PART_NUMBER" type="text" id="PART_NUMBER" value="<%=PART_NUMBER%>"  readonly="true" style="background-color:#666666"></td>
      <td >Inner Pot</td>
      <td colspan="2" ><input name="BATCHNOB" type="text" id="BATCHNOB" value="<%=BATCHNOB%>" size="30" ></td>
      <td >&nbsp;</td>
    </tr>
    <tr>
      <td >Shift</td>
      <td ><select name="Shift" id="Shift">
       
       <option value="A">A</option>
        <option value="B">B</option>
        <option value="C">C</option>
        <option value="D">D</option>
      </select></td>
      <td >&nbsp;</td>
      <td width="104" >&nbsp;</td>
      <td width="255" >&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
     <tr>
      <td >Date</td>
      <td ><input name="pack_Time" type="text" id="pack_Time" value="<%=formatdate(Now,"yyyymmdd")%>" readonly style="background-color:#666666"></td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
    </tr>
    <tr>
      <td >Qty</td>
      <td ><input name="Qty" type="text" id="Qty" value="<%=PACKED_QTY%>" ></td>
      <td >Equipment NO</td>
      <td colspan="3" ><input name="Equipment_NO" type="text" id="Equipment_NO" value="<%=Equipment_NO%>" ></td>
    </tr>
	<tr>
      <td >Remark</td>
      <td ><input name="Remark" type="text" id="Remark" value=<%=Remark%> ></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    	<tr>
      <td >JOB Number</td>
      <td ><input name="JOB_NUMBER" type="text" id="JOB_NUMBER" value=<%=JOB_NUMBER%> ></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center">
	
	  <input type="hidden" id="computername" name="computername" value="<%=computername%>">
      <input type="hidden" id="PACKED_QTY" name="PACKED_QTY" value="<%=PACKED_QTY%>">
	  <input name="btnSubmit" type="button" id="btnSubmit"  value="Print 打印" onClick="Print();">  
	  &nbsp;
	  <input name="btnclose" type="button"  id="btnclose"  value="Close 关闭" onClick="javascript:window.close();">
      &nbsp;  
          
	  </td>
    </tr>
   <tr align="center"><td colspan="6">&nbsp;</td></tr>
  </form>
</table>
</body>

</html>

<!--#include virtual="/WOCF/BOCF_Close.asp" -->