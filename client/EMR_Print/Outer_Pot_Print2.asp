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
			window.alert("�����빤���ţ�");
			document.form1.JOB_NUMBER.focus();
			return;
		}
		  
		   if(document.form1.Inspector.value=="")
		{
			window.alert("���������ˣ�");
			document.form1.Inspector.focus();
			return;
		}
		  
		  
		     if(document.form1.BATCHNOA.value=="")
		{
			window.alert("Batch no ����Ϊ�գ�");
			document.form1.BATCHNOA.focus();
			return;
		} 
		
		    if(document.form1.BATCHNOB.value=="")
		{
			window.alert("Batch no ����Ϊ�գ�");
			document.form1.BATCHNOB.focus();
			return;
		} 
		
		    if(document.form1.BATCHNOC.value=="")
		{
			window.alert("Batch no ����Ϊ�գ�");
			document.form1.BATCHNOC.focus();
			return;
		} 
		  
		  
		    if(document.form1.BATCHNOD.value=="")
		{
			window.alert("Batch no ����Ϊ�գ�");
			document.form1.BATCHNOD.focus();
			return;
		}   
		  
		  
		  
		  
		  
		    if(document.form1.PRODTIMEA.value=="")
		{
			window.alert("�������ڲ���Ϊ�գ�");
			document.form1.PRODTIMEA.focus();
			return;
		} 
		
		
		  if(document.form1.PRODTIMEB.value=="")
		{
			window.alert("�������ڲ���Ϊ�գ�");
			document.form1.PRODTIMEB.focus();
			return;
		} 
		  if(document.form1.PRODTIMEC.value=="")
		{
			window.alert("�������ڲ���Ϊ�գ�");
			document.form1.PRODTIMEC.focus();
			return;
		} 
		  if(document.form1.PRODTIMED.value=="")
		{
			window.alert("�������ڲ���Ϊ�գ�");
			document.form1.PRODTIMED.focus();
			return;
		} 
		
		
		
		
		
		  
		  if(document.form1.Qty.value>document.form1.PACKED_QTY.value)
		  {
			  window.alert("�������ܴ���Ĭ����ֵ");
			document.form1.Qty.focus();
			return;
			  
		}
		  
		  
		  
		 if (!document.form1.Equipment_NO.value)
   		 {
    		alert("����������ţ�");
			document.form1.Equipment_NO.focus();
			return false;	
		
		}		
		
		
		else{
			
			form1.action="Outer_Pot_print2.asp?Action=2";
			form1.submit();
		}
		 
		 
		 
		 
		 
	
	 }
	 
	 
</script>
<%
computername=trim(request("computername"))
part_NUMBER=trim(request("PART_NUMBER"))

packed_USER=trim(request("PACKED_USER"))


	

 
		    sql="select * from Product_model where item_name='"&PART_NUMBER&"'"
			set rs1=server.createobject("adodb.recordset")
			rs1.open SQL,conn,1,3
			if not rs1.eof then	
			PRODUCTION_NAME=rs1("Description")
			PACKED_QTY=rs1("Box_size")
			SMALL_pack=rs1("SMALL_pack")
			CCL=rs1("CCL")
				if isnull(CCL) then
				word="�������߱�CCL����"
				response.Redirect("Pot_print.asp?word="&word)
				end if
				if isnull(SMALL_pack) then
				word="��������С��װ������"
				response.Redirect("Pot_print.asp?word="&word)
				end if
					
		
			else
			word="�ϺŲ����ڣ���������ȷ���Ϻ�"		
			response.Redirect("Pot_print.asp?word="&word)
			end if
			rs1.close
			set rs4=nothing
			Action=request("Action")

		 
		' end if
		 
			If Action="2" then
			conn.begintrans
			set rsAsyn=server.CreateObject("adodb.recordset")
				
				
				PRODUCTION_NAME=trim(request("PRODUCTION_NAME"))				
								
				boxId=left(trim(request("Box_id")),12)
				Shift=trim(request("Shift"))
				pack_Time=trim(request("pack_Time"))
				qty=trim(request("qty"))
				POT=trim(request("POT"))
				TOP_Ring=trim(request("TOP_Ring"))				
				Magent_L=trim(request("Magent_L"))
				Magent_S=trim(request("Magent_S"))
				Equipment_NO=trim(request("Equipment_NO"))
				Inspector=trim(request("Inspector"))
				Remark=trim(request("Remark"))
				BATCHNOA=trim(request("BATCHNOA"))
				BATCHNOAA=mid(BATCHNOA,8,10)
				
				BATCHNOB=trim(request("BATCHNOB"))
                BATCHNOBB=right(BATCHNOB,13)
			
				
				BATCHNOC=trim(request("BATCHNOC"))
				BATCHNOCC=right(BATCHNOC,7)
				
				
				BATCHNOD=trim(request("BATCHNOD"))
				BATCHNODD=right(BATCHNOD,7)
				
				PRODTIMEA=trim(request("PRODTIMEA"))
				PRODTIMEB=trim(request("PRODTIMEB"))
				PRODTIMEC=trim(request("PRODTIMEC"))
				PRODTIMED=trim(request("PRODTIMED"))
				
				
				JOB_NUMBER=trim(request("JOB_NUMBER"))
								
					SQLH="SELECT * FROM tbl_MES_LotMaster  where WipEntityName='"&JOB_NUMBER&"'"
			
				rsTicket.open SQLH,connTicket,1,3
					if  rsTicket.bof and rsTicket.eof then
					  '  word="�˹��������ڣ���ȷ�Ϲ�����"
						'response.Redirect("Pot_print.asp?word="&word)
					END IF	
				rsTicket.close				
				'CCL=mid(JOB_NUMBER,3,4)
				
					if boxId="" then
							'����box id						
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


          
				'���ô�ӡ����
				sql="SELECT PRINTER_NAME FROM COMPUTER_PRINTER_MAPPING WHERE COMPUTER_NAME='"+UCase(request("computername"))+"'"
   
				rsAsyn.open sql,conn,1,3
				if not rsAsyn.eof then
					printName = rsAsyn("PRINTER_NAME")
				end if
				rsAsyn.close
				if printName="" then
					word="Please contact engiener to set the printer for this machine.|����ϵ����ʦΪ�˻����趨��ȷ�ı�ǩ��ӡ����"
					
					response.Redirect("Pot_print.asp?word="&word)
				else
				  
   
					           
					set PrintCtl=server.createobject("PrintClass.PrintCtl")
					


					
					JSQ=int(cint(qty)/cint(SMALL_pack))
					'response.Write(JSQ)
					'response.Write("<BR>")
				    LastQty=cint(cint(qty)-cint(SMALL_pack)*JSQ)
					'response.Write(LastQty)
					'response.Write("<BR>")
					'response.End()
					if JSQ>0 then
						
						for i=1 to JSQ 
						pack_TimeL=formatdate(now,"yyyymmdd")
						boxIdN=boxId&"-"&i 
						'response.Write(boxIdN)
						'response.Write("<BR>")
						'rtnCode=PrintCtl.PrintKEBPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxIdN,Shift,pack_TimeL,SMALL_pack,POT,TOP_Plate,Magent,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark)
					    rtnCode=PrintCtl.PrintKEBOUTERPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxIdN,Shift,pack_TimeL,SMALL_pack,POT,TOP_Ring,Magent_L,Magent_S,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark,BATCHNOAA,BATCHNOBB,BATCHNOCC,BATCHNODD,PRODTIMEA,PRODTIMEB,PRODTIMEC,PRODTIMED)

						TimeDelaySeconds(1)
						'rtnCode="OK"
						if rtnCode<>"OK" then
							word="Label print error.��ӡ��ǩ��������.|" & rtnCode		
							response.Redirect("Pot_print.asp?word="&word)
						end if
						
						next
					end if	
					
				if LastQty>0 then
				boxIdN=boxId&"-"&i
					rtnCode=PrintCtl.PrintKEBOUTERPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxIdN,Shift,pack_TimeL,LastQty,POT,TOP_Ring,Magent_L,Magent_S,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark,BATCHNOAA,BATCHNOBB,BATCHNOCC,BATCHNODD,PRODTIMEA,PRODTIMEB,PRODTIMEC,PRODTIMED)
					if rtnCode<>"OK" then
				            word="Label print error.��ӡ��ǩ��������.|" & rtnCode		
							response.Redirect("Pot_print.asp?word="&word)		
				    end if
				end if
					
					
					'rtnCode=PrintCtl.PrintKEBOUTERPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxId,Shift,pack_Time,qty,POT,TOP_Ring,Magent_L,Magent_S,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark)
				   	rtnCode=PrintCtl.PrintKEBOUTERPOTLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxId,Shift,pack_Time,qty,POT,TOP_Ring,Magent_L,Magent_S,Equipment_NO,PACKED_USER,InSpector,JOB_NUMBER,Remark,BATCHNOAA,BATCHNOBB,BATCHNOCC,BATCHNODD,PRODTIMEA,PRODTIMEB,PRODTIMEC,PRODTIMED)

				    if rtnCode="OK" then
					PACK_TIME=now()
					SUPPLIERA="POT-"&POT
					SUPPLIERB="TOP-Ring-"&TOP_Ring
					SUPPLIERC="Magent-L-"&Magent_L 
					SUPPLIERD="Magent-S-"&Magent_S
					'sql="insert into EMR_PACK_DETAIL  (PACK_USER,PART_NUMBER,PRODUCTION_NAME,BOX_ID,SHIFT,PACK_TIME,PACK_QTY,SUPPLIERA,SUPPLIERB,SUPPLIERC,EQUIPMENT_NO,INSPECTOR,REMARK,JOB_NUMBER,SUPPLIERD,BATCHNOA,BATCHNOB,BATCHNOC,BATCHNOD,PRODTIMEA,PRODTIMEB,PRODTIMEC,PRODTIMED) VALUES  ('"+packed_USER+"','"+part_NUMBER+"','"+PRODUCTION_NAME+"','"+boxId+"','"+Shift+"','"&PACK_TIME&"','"+qty+"','"+SUPPLIERA+"','"+SUPPLIERB+"','"+SUPPLIERC+"','"+EQUIPMENT_NO+"','"+INSPECTOR+"','"+REMARK+"','"+JOB_NUMBER+"','"+SUPPLIERD+"','"+BATCHNOA+"','"+BATCHNOB+"','"+BATCHNOC+"','"+BATCHNOD+"','"+PRODTIMEA+",'"+PRODTIMEB+"','"+PRODTIMEC+"','"+PRODTIMED+"')"
					sql="insert into EMR_PACK_DETAIL  (PACK_USER,PART_NUMBER,PRODUCTION_NAME,BOX_ID,SHIFT,PACK_TIME,PACK_QTY,SUPPLIERA,SUPPLIERB,SUPPLIERC,EQUIPMENT_NO,INSPECTOR,REMARK,JOB_NUMBER,SUPPLIERD,BATCHNOA,BATCHNOB,BATCHNOC,BATCHNOD,PRODTIMEA,PRODTIMEB,PRODTIMEC,PRODTIMED,JSQ) VALUES  ('"+packed_USER+"','"+part_NUMBER+"','"+PRODUCTION_NAME+"','"+boxId+"','"+Shift+"','"&PACK_TIME&"','"+qty+"','"+SUPPLIERA+"','"+SUPPLIERB+"','"+SUPPLIERC+"','"+EQUIPMENT_NO+"','"+INSPECTOR+"','"+REMARK+"','"+JOB_NUMBER+"','"+SUPPLIERD+"','"+BATCHNOA+"','"+BATCHNOB+"','"+BATCHNOC+"','"+BATCHNOD+"','"+PRODTIMEA+"','"+PRODTIMEB+"','"+PRODTIMEC+"','"+PRODTIMED+"','"+qty+"')"

					conn.execute(sql)
					else
						word="Label print error.��ӡ��ǩ��������.|" & rtnCode		
					end if


				  
					
				end if
			if conn.errors.count>0 then
				
				conn.rollbacktrans
			else 
			    conn.commitTrans     'ִ�������ύ
			  ' response.Redirect("Pot_print.asp")
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
      <td height="20" colspan="6" class="t-t-DarkBlue"  align="center">(Outer Pot Assy Print)</td>
    </tr>
	
    <tr>
      <td width="152" rowspan="2" >Operator Code </td>
      <td width="328" rowspan="2" ><input name="PACKED_USER" type="text" id="PACKED_USER" value=<%=PACKED_USER%>  readonly="true" style="background-color:#666666"></td>
      <td width="121" >Inspector</td>
      <td colspan="3" ><label for="select">
        <input name="Inspector" type="text" id="Inspector" value=<%=Inspector%> >
      </label></td>
    </tr>
    <tr>
      <td height="25" >Description</td>
      <td width="133" >Supplier </td>
      <td width="255" >Batch No</td>
      <td width="197" >Date NO</td>
    </tr>
    <tr>
      <td >Production</td>
      <td ><input name="PRODUCTION_NAME" type="text" id="PRODUCTION_NAME" value="<%=PRODUCTION_NAME%>" size="40" readonly style="background-color:#666666" ></td>
      <td >POT</td>
      <td ><select name="POT" id="POT">

<option value="HZ">HZ</option>
<option value="ChaoYa">ChaoYa</option>
<option value="LM">LM</option>
<option value="KEM2">KEM2</option>
<option value="Yabets">Yabets</option>
<option value="YiDong">YiDong</option>
<option value="NA">NA</option>
      </select></td>
      <td ><input name="BATCHNOA" type="text" id="BATCHNOA" value="<%=BATCHNOA%>" size="30" ></td>
      <td ><input name="PRODTIMEA" type="text" id="PRODTIMEA" value="<%=PRODTIMEA%>" size="20" ></td>
    </tr>
    <tr>
      <td >12NC</td>
      <td ><input name="PART_NUMBER" type="text" id="PART_NUMBER" value="<%=PART_NUMBER%>"  readonly="true" style="background-color:#666666"></td>
      <td >TOP-Ring</td>
      <td ><select name="TOP_Ring" id="TOP_Ring">
      <option value="HZ">HZ</option>
       <option value="ChaoYa">ChaoYa</option>
        <option value="YiDong">YiDong</option>
        <option value="Yabets">Yabets</option>
        <option value="KEM2">KEM2</option>
        <option value="NA">NA</option>
      </select></td>
      <td ><input name="BATCHNOB" type="text" id="BATCHNOB" value="<%=BATCHNOB%>" size="30" ></td>
      <td ><input name="PRODTIMEB" type="text" id="PRODTIMEB" value="<%=PRODTIMEB%>" size="20" ></td>
    </tr>
    <tr>
      <td >Shift</td>
      <td ><select name="Shift" id="Shift">
       
       <option value="A">A</option>
        <option value="B">B</option>
        <option value="C">C</option>
        <option value="D">D</option>
      </select></td>
      <td >Magent-L</td>
      <td ><select name="Magent_L" id="Magent_L">
        <option value="TongLi">TongLi</option>
        <option value="BJMT">BJMT</option>
        <option value="Quadrant">Quadrant</option>
        <option value="SLMG">SLMG</option>
        <option value="NA">NA</option>
        
      </select></td>
      <td ><input name="BATCHNOC" type="text" id="BATCHNOC" value="<%=BATCHNOC%>" size="30" ></td>
      <td ><input name="PRODTIMEC" type="text" id="PRODTIMEC" value="<%=PRODTIMEC%>" size="20" ></td>
    </tr>
     <tr>
      <td >Date</td>
      <td ><input name="pack_Time" type="text" id="pack_Time" value="<%=formatdate(Now,"yyyymmdd")%>" readonly style="background-color:#666666"></td>
      <td >Magent-S</td>
      <td ><select name="Magent_S" id="Magent_S">
        <option value="TongLi">TongLi</option>
        <option value="BJMT">BJMT</option>
        <option value="Quadrant">Quadrant</option>
        <option value="SLMG">SLMG</option>
        <option value="NA">NA</option>
      </select></td>
      <td ><input name="BATCHNOD" type="text" id="BATCHNOD" value="<%=BATCHNOD%>" size="30" ></td>
      <td ><input name="PRODTIMED" type="text" id="PRODTIMED" value="<%=PRODTIMED%>" size="20" ></td>
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
	  <input name="btnSubmit" type="button" id="btnSubmit"  value="Print ��ӡ" onClick="Print();">  
	  &nbsp;
	  <input name="btnclose" type="button"  id="btnclose"  value="Close �ر�" onClick="javascript:window.close();">
      &nbsp;  
          
	  </td>
    </tr>
   <tr align="center"><td colspan="6">&nbsp;</td></tr>
  </form>
</table>
</body>

</html>

<!--#include virtual="/WOCF/BOCF_Close.asp" -->