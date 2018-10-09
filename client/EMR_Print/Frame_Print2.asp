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
		 	  
		   if(document.form1.Inspector.value=="")
		{
			window.alert("请输入检查人！");
			document.form1.Inspector.focus();
			return;
		}
		  
		  
		  
		
		  
	else{
			
			form1.action="Frame_print2.asp?Action=2";
			form1.submit();
		}
		 
	 }
	 
	 
</script>
<%
computername=trim(request("computername"))
inputJob=trim(request("Job_number"))
packed_USER=trim(request("PACKED_USER"))

if instr(inputJob,"-")=0 then
	response.Redirect("Frame_print.asp?word=Job Number is error, please scan sub job！<br>工单错误，请扫描子工单！")
end if
ajobnumber=split(inputJob,"-")
strSheetNo=ajobnumber(ubound(ajobnumber))

if len(strSheetNo)>3 then
	response.Redirect("Frame_print.asp?word=Job Number is error, please scan sub job！<br>细分工单单编号超过3位错误，请扫描子工单！")
end if

if isnumeric(strSheetNo) then
	sheetnumber = cint(strSheetNo)
else
	response.Redirect("Frame_print.asp?word=Job Number is error, please re-scan sub job！<br>工单错误，请重新扫描子工单！")
end if
jobnumber=left(inputJob,len(inputJob)-4)
 
sql="select * from job where job_number='"&jobnumber&"' and sheet_number='"&sheetnumber&"'"
'response.Write(sql)
'response.End()
rs.open sql,conn,1,3
if not rs.eof then
	PART_NUMBER=rs("PART_NUMBER_TAG")
	LINE_NAME=rs("LINE_NAME")
	qty=rs("JOB_GOOD_QUANTITY")
else
	word="工单不存在，是否已经开站？"		
	response.Redirect("Frame_print.asp?word="&word)
end if
rs.close
	
	
            sql="select * from Product_model where item_name='"&PART_NUMBER&"'"
			set rs1=server.createobject("adodb.recordset")
			rs1.open SQL,conn,1,3
			if not rs1.eof then	
			PRODUCTION_NAME=rs1("Description")			
            else
			word="料号不存在，请输入正确的料号"		
			response.Redirect("Frame_print.asp?word="&word)
			end if
			rs1.close
			set rs1=nothing
			Action=request("Action")		
			
			
If Action="2" then
			    conn.begintrans
			    set rsAsyn=server.CreateObject("adodb.recordset")
				PRODUCTION_NAME=trim(request("PRODUCTION_NAME"))				
				boxId=left(trim(request("Box_id")),12)
				Shift=trim(request("Shift"))
				pack_Time=trim(request("pack_Time"))
				qty=trim(request("qty"))
				Equipment_NO=trim(request("Equipment"))
				Inspector=trim(request("Inspector"))
				Remark=trim(request("Remark"))							
				CCL=Equipment_NO
				
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
	
					rtnCode=PrintCtl.PrintKEBMEMASSYLabel(printName,PRODUCTION_NAME,PART_NUMBER,boxId,Shift,pack_Time,qty,Equipment_NO,PACKED_USER,InSpector,inputJob,Remark)

						
						rtnCode="OK"
						if rtnCode<>"OK" then
							word="Label print error.打印标签发生错误.|" & rtnCode		
							response.Redirect("Pot_print.asp?word="&word)
						end if
						
				
				end if
					
					
					

				    if rtnCode="OK" then
					PACK_TIME=now()
					sql="insert into EMR_PACK_DETAIL  (PACK_USER,PART_NUMBER,PRODUCTION_NAME,BOX_ID,SHIFT,PACK_TIME,PACK_QTY,EQUIPMENT_NO,INSPECTOR,REMARK,JOB_NUMBER,JSQ) VALUES  ('"+packed_USER+"','"+part_NUMBER+"','"+PRODUCTION_NAME+"','"+boxId+"','"+Shift+"','"&PACK_TIME&"','"+qty+"','"+EQUIPMENT_NO+"','"+INSPECTOR+"','"+REMARK+"','"+inputJob+"','"+qty+"')"

					conn.execute(sql)
					else
						word="Label print error.打印标签发生错误.|" & rtnCode		
					end if
				
				
			if conn.errors.count>0 then
				
				conn.rollbacktrans
			else 
			    conn.commitTrans     '执行事务提交
			   response.Redirect("Frame_Print.asp")
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
      <td height="20" colspan="6" class="t-t-DarkBlue"  align="center">( Frame Assy Print)</td>
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
      <td height="25" >Equipment NO</td>
      <td width="133" >
      <input name="Equipment" type="text" id="Equipment" value=<%=LINE_NAME%> readonly style="background-color:#666666" ></td>
      <td colspan="2" >&nbsp;</td>
    </tr>
    <tr>
      <td >Production</td>
      <td ><input name="PRODUCTION_NAME" type="text" id="PRODUCTION_NAME" value="<%=PRODUCTION_NAME%>" size="40" readonly style="background-color:#666666" ></td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td colspan="2" >&nbsp;</td>
    </tr>
    <tr>
      <td >12NC</td>
      <td ><input name="PART_NUMBER" type="text" id="PART_NUMBER" value="<%=PART_NUMBER%>"  readonly="true" style="background-color:#666666"></td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td colspan="2" >&nbsp;</td>
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
      <td >&nbsp;</td>
      <td colspan="2" >&nbsp;</td>
    </tr>
     <tr>
      <td >Date</td>
      <td ><input name="pack_Time" type="text" id="pack_Time" value="<%=formatdate(Now,"yyyymmdd")%>" readonly style="background-color:#666666"></td>
      <td >&nbsp;</td>
      <td >&nbsp;</td>
      <td width="255" >&nbsp;</td>
      <td width="197" >&nbsp;</td>
    </tr>
    <tr>
      <td >Qty</td>
      <td ><input name="Qty" type="text" id="Qty" value="<%=qty%>" readonly style="background-color:#666666"></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
	<tr>
      <td >Remark</td>
      <td ><input name="Remark" type="text" id="Remark" value=<%=Remark%> ></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    	<tr>
      <td >JOB Number</td>
      <td ><input name="JOB_NUMBER" type="text" id="JOB_NUMBER" value=<%=inputJob%> readonly style="background-color:#666666" ></td>
      <td >&nbsp;</td>
      <td colspan="3" >&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="6" align="center">
	
	  <input type="hidden" id="computername" name="computername" value="<%=computername%>">
      
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