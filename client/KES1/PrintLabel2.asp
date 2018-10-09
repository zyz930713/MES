<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
code=request.Form("txtOperatorCode")
'clientIP=Request.ServerVariables("remote_host")
ComputerName=request("computername")
PrintName=GetPrinterName(ComputerName)

if(PrintName="")then
	word="<span align='center' style='color:red;'>Please contact engiener to set the printer for this machine.<br>请联系工程师为此机器设定正确的标签打印机。</span>"
	response.Redirect("PrintLabel.asp?word="&word)
end if 
	
SQL="select OPERATOR_NAME,AUTHORIZED_STATIONS_ID,AUTHORIZED_PARTS_ID,LOCKED,PRACTISED,PRACTISE_START_TIME,PRACTISE_END_TIME from OPERATORS where code='"&code&"'"
set rs4=server.createobject("adodb.recordset")
rs4.open SQL,conn,1,3
if rs4.eof then	
	word="<span align='center' style='color:red;'>Operator Code of "&code&" is not existed, please contact engineer.<br>"&code&" 不存在，请联系工程师。</span>"
	response.Redirect("PrintLabel.asp?word="&word)
end if
rs4.close
set rs4=nothing

	SubJob=request.Form("txtSubJobList")   
	NewJob=request.Form("newJob")
	ManJobNumber=request.Form("batchNo") 
	Action=request.querystring("Action")
	TotalCount=request.Form("TotalCount")	
	OperatorCode=request.Form("txtOperatorCode")	
	
	Set TypeLib = CreateObject("Scriptlet.TypeLib")
    Guid = TypeLib.Guid
    	
	SlappingQty=0	
	NewJob=replace(NewJob,"-",",")
	
	'check the sub job list is printted or not 2012-3-22
	ExistingGuid=""
	NewJobArr=split(NewJob,",")
	SQL=" select DISTINCT SUBJOBLIST,ID from label_print_history where job_NUMBER='"+ManJobNumber+"' AND"
	for i=0 to ubound(NewJobArr)
		SQLS=SQLS+"  (subjoblist LIKE '"+NewJobArr(i)+"-%' OR subjoblist LIKE '%-"+NewJobArr(i)+"-%' OR  subjoblist LIKE '%-"+NewJobArr(i)+"' OR  subjoblist = '"+NewJobArr(i)+"') OR"
	next 
	if(SQLS<>"") then
		SQL=SQL+"("+left(SQLS,len(SQLS)-2)+")"
	end if 
	
	SQL=SQL+" and transactiontype='-1'"	
	rs.open SQL,conn,1,3
	if(rs.recordcount=1)then
		ExistingGuid=rs("ID")
		ExistingSubJobList=rs("SUBJOBLIST")
		ExistingSubJobArr=split(ExistingSubJobList,"-")
		if(ubound(ExistingSubJobArr)<>ubound(NewJobArr))then
			word="<span align='center' style='color:red;'>此批子工单已经打印过，且和现在的子工单的个数不一样:"+ExistingSubJobList+"</span>"
		end if 
	elseif(rs.recordcount>1)then
		str=""
		while not rs.eof
			str=str+rs("SUBJOBLIST")+"<br>"
			rs.movenext
		wend
		word="<span align='center' style='color:red;'>此批子工单已经打印过，之前是分开打印，现在不能和在一起打印:"+str+"</span>"		
	end if
	rs.close
	if word <> "" then
		response.Redirect("PrintLabel.asp?word="&word)
	end if
	
	'check OQC
	sql="select batchnolist from job_oqc a where exists(select 1 from label_print_history where batchno=a.batchnolist and id='"&ExistingGuid&"')"
	rs.open sql,conn,1,3
	if not rs.eof then
		batchNo=split(rs("batchnolist"),"--")(0)
		isReprint="1"
	end if
	rs.close
	
	'Check finish main line
	SQL0="SELECT SHEET_NUMBER,CLOSE_TIME FROM JOB WHERE JOB_NUMBER='"&ManJobNumber&"' AND SHEET_NUMBER IN ("&NewJob&") "
	rs.open SQL0,conn,1,3
	UnFinishList=""
	existList=""
	while not rs.eof
		if isnull(rs("CLOSE_TIME")) then
			UnFinishList=UnFinishList+cstr(rs("SHEET_NUMBER"))+","
		end if
		existList=existList+cstr(rs("SHEET_NUMBER"))+","
		rs.movenext
	wend
	rs.close
	if UnFinishList<>"" then
		word="<span align='center' style='color:red;'>子工单："+UnFinishList+"还没有完成主线,不能打标签！</span>"
		response.Redirect("PrintLabel.asp?word="&word)
	end if		

	'Check 子工单的重复性
	ArrSubJobList=split(newjob,",")
	for i=0 to ubound(ArrSubJobList)
		if instr(existList,ArrSubJobList(i))=0 then
			word="<span align='center' style='color:red;'>子工单："+ArrSubJobList(i)+"不存在,不能打印标签！</span>"
			response.Redirect("PrintLabel.asp?word="&word)
		end if	
	next 

	'Rework , Change model, Scrap, Readjust Qty
	SQL1="SELECT SUM(JD.DEFECT_QUANTITY) AS QTY, DC.TRANSACTION_TYPE "
	SQL1=SQL1+" FROM JOB_DEFECTCODES JD , DEFECTCODE DC"
	SQL1=SQL1+ " WHERE JD.DEFECT_CODE_ID=DC.NID(+)"
	SQL1=SQL1+" AND JD.JOB_NUMBER='"&ManJobNumber&"' AND JD.SHEET_NUMBER in ("&NewJob&") AND DC.TRANSACTION_TYPE IN ('0','1','2','3','4','5')"
	SQL1=SQL1+" GROUP BY DC.TRANSACTION_TYPE"
	SQL1=SQL1+" ORDER BY DC.TRANSACTION_TYPE"
	
	 
	set rs1=server.createobject("adodb.recordset")
	rs1.open SQL1,conn,1,3

	'Normal Qty		
	SQL2="SELECT  SUM(J.JOB_GOOD_QUANTITY) AS QTY"
	SQL2=SQL2+" FROM job J "
	SQL2=SQL2+ " WHERE J.JOB_NUMBER='"&ManJobNumber&"' AND J.SHEET_NUMBER in ("&NewJob&") "	
	rs.open SQL2,conn,1,3
	if not rs.eof then
		normalQty=rs("QTY")  '得到良品数量
	end if
	rs.close
	
	'Slapping Qty
	SQL3=" select  DECODE(SUM(A.ACTION_VALUE),NULL,0,SUM(A.ACTION_VALUE)) AS QTY from job_actions a,action b "
	SQL3=SQL3+" where a.action_id=b.nid and job_number='"&ManJobNumber&"' and A.SHEET_NUMBER  in ("&NewJob&")  and b.mother_action_id='AN00000260'"	 
 	
	rs.open SQL3,conn,1,3	 
	if rs.recordcount<>0 then
		SlappingQty=rs("QTY")
	end if
	rs.close
	
If Action="2" then
	if NewJob<>"" then
		set PrintCtl=server.createobject("PrintClass.PrintCtl")
		ReturnCode=""
		if isReprint <> "1" then
			SQL="Select batchno From LABEL_PRINT_HISTORY where Job_Number='"&ManJobNumber&"' order by printtime desc "
			rs.open SQL,conn,1,3
			if  rs.recordcount=0 then
				batchNo=ManJobNumber&"-1"
			else
				aryBatchNo = split(rs("batchno"),"-")
				if Ubound(aryBatchNo)>=2 then
					if isNumeric(aryBatchNo(1)) then
						maxnumber = aryBatchNo(1)
					else 
						maxnumber = aryBatchNo(2)
					end if
				end if						
				batchNo=ManJobNumber&"-"&(cint(maxnumber)+1)
			end if
			rs.close			
		end if
		
		TransactionType=""
		Qty=""
		Configuration=""
		SQL="SELECT DISTINCT PART_NUMBER_TAG FROM JOB J WHERE J.JOB_NUMBER='"&ManJobNumber&"'  and J.SHEET_NUMBER  in ("&NewJob&") "
		rs.open SQL,conn,1,3
		IF rs.recordcount>0 then
			Configuration=rs("PART_NUMBER_TAG")
		end if
		if isReprint <> "1" then
		'add for update history 
			IF(ExistingGuid<>"")Then
				SQL="delete from label_print_history where id='"+ExistingGuid+"' "
				set rshisotrydelete1=server.createobject("adodb.recordset")
				rshisotrydelete1.open SQL,conn,1,3
				SQL="update JOB_RETEST_QA_HISTORY set IS_DELETE='1' where GUID='"+ExistingGuid+"'"
				set rshisotrydelete2=server.createobject("adodb.recordset")
				rshisotrydelete2.open SQL,conn,1,3
			end if 
		end if	
		rs.close
 
		'END ADD
		for i=-1 to 5
			TransactionType=request.Form("txtTransactionType"&cstr(i))
			SystemQty=request.Form("txtQty"&cstr(i))
			Qty=request.Form("txtActualQty"&cstr(i))
			Remark=request.Form("txtComments"&cstr(i))
			NewJob=replace(NewJob,",","-")
 
			if(TransactionType="-1") then
				if(SystemQty>0) then
					Qty=cint(Qty)
					SubJobArr=split(NewJob,"-")
					for mm=0 to ubound(SubJobArr)
						if cint(SubJobArr(mm))>=500 then
							isRework=true
						end if 
					next 
					if isRework=true then
							RetestRate="100%"
							SampleSize=Qty
					else
						'add for Retest Rate
						SQL="SELECT * FROM General_Setting WHERE MODELNAME='"+Configuration+"'"
						rs.open SQL,conn,1,3
						if rs.recordcount=0 or isnull(rs("RetestRatio"))  then
							word="<span align='center' style='color:red;'>此型号没有设定General Setting.</span>"
							response.Redirect("PrintLabel.asp?word="&word)
						else
							Ratio=rs("RetestRatio")
							RetestRate=cstr(cdbl(Ratio)*100)+"%"
							if (Ratio=0) then
								SampleSize=cdbl(cstr(cint(Qty)))
							else
								'SampleSize=cdbl(cstr(cint(Qty)*cdbl(Ratio)))
								SampleSize=round(cdbl(cstr(cint(Qty)*cdbl(Ratio))))
							end if
						end if
						rs.close
					end if
					ReturnCode=PrintCtl.PrintNormal(PrintName, 1,batchNo+"-"+TransactionType, NewJob, Qty,Configuration,SampleSize,RetestRate)					
					'ReturnCode="OK"
					if(ReturnCode="OK") then
						if isReprint <> "1" then
						SQL="SELECT * FROM LABEL_PRINT_HISTORY WHERE 1=2"						
						rs.open SQL,conn,1,3
						rs.addnew
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
 
							rs("QTY")=SystemQty
							rs("SUBJOBLIST")=replace(NewJob,"'","")
							rs("ACTUALQTY")=Qty
							rs("OPERATOR")=OperatorCode
							rs("REMARK")=Remark
							rs("PRINTTIME")=NOW()
							rs("TRANSACTIONTYPE")=TransactionType
							rs("id")=Guid
							rs("RETESTRATIO")=RetestRate
							rs("RETESTQTY")=SampleSize
						rs.update
						rs.close
						
						SQL="SELECT * FROM JOB_RETEST_QA_HISTORY WHERE 1=2"
						rs.open SQL,conn,1,3
						rs.addnew
							rs("GUID")=Guid
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("STATION_NAME")="RETEST"
							rs("MODEL_NAME")=Configuration
							rs("STATION_IN_DATETIME")=now()
							rs("STATION_IN_QTY")=Qty
							rs("STATION_IN_OPERATOR")=OperatorCode
							rs("IS_SPLAPPING")="0"
							rs("IS_DELETE")="0"
							rs("IS_FINISHED")="0"
						rs.update
						rs.close
						end if
						response.write("<script>window.alert('Print Normal label successfully!')</script>")
					else
						response.write("<script>window.alert('Print Normal label failed!')</script>")
					end if
				end if
			end if
			
			if(TransactionType="1") then				
				ReturnCode=""
				if(SystemQty>0) then
					ReturnCode=PrintCtl.PrintRework(PrintName, 1,batchNo+"-"+TransactionType, NewJob, Qty,Configuration)
					'ReturnCode="OK"
					if(ReturnCode="OK") then
						if isReprint <> "1" then
						SQL="SELECT * FROM LABEL_PRINT_HISTORY WHERE 1=2"						
						rs.open SQL,conn,1,3
						rs.addnew
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("QTY")=SystemQty
							rs("SUBJOBLIST")=replace(NewJob,"'","")
							rs("ACTUALQTY")=Qty
							rs("OPERATOR")=OperatorCode
							rs("REMARK")=Remark
							rs("PRINTTIME")=NOW()
							rs("TRANSACTIONTYPE")=TransactionType
							rs("id")=Guid
						rs.update
						rs.close
						end if
						response.write("<script>window.alert('Print Rework label successfully!')</script>")
					else
						response.write("<script>window.alert('Print Rework label failed!')</script>")
					end if
				end if
			end if
			
			if(TransactionType="2") then
				ReturnCode=""
				if(SystemQty>0) then
					ReturnCode=PrintCtl.PrintScrap(PrintName, 1,batchNo+"-"+TransactionType, NewJob, Qty,Configuration)
					'ReturnCode="OK"
					if(ReturnCode="OK") then
						if isReprint <> "1" then
						SQL="SELECT * FROM LABEL_PRINT_HISTORY WHERE 1=2"
						rs.open SQL,conn,1,3
						rs.addnew
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("QTY")=SystemQty
							rs("SUBJOBLIST")=replace(NewJob,"'","")
							rs("ACTUALQTY")=Qty
							rs("OPERATOR")=OperatorCode
							rs("REMARK")=Remark
							rs("PRINTTIME")=NOW()
							rs("TRANSACTIONTYPE")=TransactionType
							rs("id")=Guid
						rs.update
						rs.close
						end if
						response.write("<script>window.alert('Print Scrap label successfully!')</script>")
					else
						response.write("<script>window.alert('Print Scrap label failed!')</script>")
					end if
				end if
			end if
			
			
			if(TransactionType="3") then
				ReturnCode=""
				if(SystemQty>0) then
					ReturnCode=PrintCtl.PrintReAdjust(PrintName, 1,batchNo+"-"+TransactionType, NewJob, Qty,Configuration)
					'ReturnCode="OK"
					if(ReturnCode="OK") then
						if isReprint <> "1" then
						SQL="SELECT * FROM LABEL_PRINT_HISTORY WHERE 1=2"
						rs.open SQL,conn,1,3
						rs.addnew
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("QTY")=SystemQty
							rs("SUBJOBLIST")=replace(NewJob,"'","")
							rs("ACTUALQTY")=Qty
							rs("OPERATOR")=OperatorCode
							rs("REMARK")=Remark
							rs("PRINTTIME")=NOW()
							rs("TRANSACTIONTYPE")=TransactionType
							rs("id")=Guid
						rs.update
						rs.close
						end if
						response.write("<script>window.alert('Print Scrap label successfully!')</script>")
					else
						response.write("<script>window.alert('Print Scrap label failed!')</script>")
					end if
				end if
			end if
			
			
			if(TransactionType="4") then
				ReturnCode=""
				if(SystemQty>0) then
					ReturnCode=PrintCtl.PrintOthers(PrintName, 1,batchNo+"-"+TransactionType, NewJob, Qty,Configuration)
					'ReturnCode="OK"
					if(ReturnCode="OK") then
						if isReprint <> "1" then
						SQL="SELECT * FROM LABEL_PRINT_HISTORY WHERE 1=2"
						rs.open SQL,conn,1,3
						rs.addnew
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("QTY")=SystemQty
							rs("SUBJOBLIST")=replace(NewJob,"'","")
							rs("ACTUALQTY")=Qty
							rs("OPERATOR")=OperatorCode
							rs("REMARK")=Remark
							rs("PRINTTIME")=NOW()
							rs("TRANSACTIONTYPE")=TransactionType
							rs("id")=Guid
						rs.update
						rs.close
						end if
						response.write("<script>window.alert('Print Change Model label successfully!');</script>")
					else
						response.write("<script>window.alert('Print Change Model label failed!')</script>")
					end if
				end if
			end if
						
			if(TransactionType="5") then
				ReturnCode=""
				if(SystemQty>0) then
					SQL="select * from offline_store where transactiontype='5' and job_number='"+ManJobNumber+"'"
					rs.open SQL,conn,1,3
					LaserJobNumber=""
					if rs.recordcount=0 then
						LaserJobNumber=ManJobNumber+"-600"
					else
						LaserJobNumber=ManJobNumber+"-"+cstr(600+cint(rs.recordcount))
					end if
					rs.close
					
					ReturnCode=PrintCtl.PrintSlapping(PrintName, 1,batchNo+"-"+TransactionType, NewJob, Qty,Configuration,LaserJobNumber)
					'ReturnCode="OK"
					if(ReturnCode="OK") then
						if isReprint <> "1" then
						SQL="SELECT * FROM LABEL_PRINT_HISTORY WHERE 1=2"
						rs.open SQL,conn,1,3
						rs.addnew
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("QTY")=SystemQty
							rs("SUBJOBLIST")=replace(NewJob,"'","")
							rs("ACTUALQTY")=Qty
							rs("OPERATOR")=OperatorCode
							rs("REMARK")=Remark
							rs("PRINTTIME")=NOW()
							rs("TRANSACTIONTYPE")=TransactionType
							rs("id")=Guid
							rs("RETESTRATIO")="100%"
							rs("RETESTQTY")=Qty
						rs.update
						rs.close
						
						SQL="SELECT * FROM offline_store WHERE 1=2"
						rs.open SQL,conn,1,3
						rs.addnew
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("QTY")=SystemQty
							rs("SUBJOBLIST")=replace(NewJob,"'","")
							rs("ACTUALQTY")=Qty
							rs("OPERATOR")=OperatorCode
							rs("REMARK")=Remark
							rs("INTIMESTAMP")=NOW()
							rs("TRANSACTIONTYPE")=TransactionType
							rs("ISPRINT")="1"
							rs("NEW_JOBNUMBER")=LaserJobNumber
						rs.update
						rs.close
						
						SQL="SELECT * FROM JOB_RETEST_QA_HISTORY WHERE 1=2"
						rs.open SQL,conn,1,3
						rs.addnew
							rs("GUID")=Guid
							rs("JOB_NUMBER")=ManJobNumber
							rs("BATCHNO")=batchNo+"-"+TransactionType
							rs("STATION_NAME")="RETEST"
							rs("MODEL_NAME")=Configuration
							rs("STATION_IN_DATETIME")=now()
							rs("STATION_IN_QTY")=Qty
							rs("STATION_IN_OPERATOR")=OperatorCode
							rs("IS_SPLAPPING")="1"
							rs("IS_DELETE")="0"
							rs("IS_FINISHED")="0"
						rs.update
						rs.close
						end if
						response.write("<script>window.alert('Print Slapping label successfully!')</script>")
					else
						response.write("<script>window.alert('Print Slapping label failed!')</script>")
					end if
				end if
			end if
		next 
	end if
end if	
%>
<%
	function GetQty(JobNumber,SheetNumber,TransactionType,Qty)
		set rsSub=server.CreateObject("adodb.recordset")
		SQL="SELECT decode(SUM(UNITQTY),null,0,SUM(UNITQTY)) FROM ENG_BORROW WHERE JOBNUMBER='"&JobNumber&"' AND SHEETNUMBER in ("&SheetNumber&") AND TRANSACTIONTYPE='" & TransactionType &"'"
		rsSub.open SQL,conn,1,3
	 	
		if Qty="" or isNull(Qty) then
			Qty=0
		end if
		if SlappingQty="" then
			SlappingQty=0
		end if
		if  rsSub.recordcount<>0 then
			if TransactionType="-1" then
				GetQty=clng(Qty)-clng(rsSub(0))-clng(SlappingQty)
			else
				GetQty=clng(Qty)-clng(rsSub(0))
			end if 
		else
			if TransactionType="-1" then
				GetQty=Qty-cint(SlappingQty)
			else
			 	GetQty=Qty
			end if
		end if
		rsSub.close
		set rsSub=nothing 
	end function
	
	
	function GetPrinterName(ComputerName)
		set rsPrinterName=server.CreateObject("adodb.recordset")
		SQL="SELECT * FROM COMPUTER_PRINTER_MAPPING WHERE UPPER(COMPUTER_NAME)='"+UCase(ComputerName)+"'"
	 
		rsPrinterName.open SQL,conn,1,3
	 	if(rsPrinterName.recordcount=0) then
			GetPrinterName=""
			exit function
		end if 
		
		if (rsPrinterName(1)="") then
			GetPrinterName=""
			exit function
		end if 
		
		GetPrinterName=rsPrinterName(1)
		rsPrinterName.close
		set rsPrinterName=nothing
	end function
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=application("SystemName")%></title>
<link href="/CSS/GeneralKES1.css" rel="stylesheet" type="text/css">
<script language="javascript"  type="text/javascript">
	 function Print()
	 {
		var ControlName="";
		var Flag=0;
	 	for(var i=-1;i<=5;i++)
		{
			ControlName1="txtTransactionType"+i.toString();
			ControlName2="txtQty"+i.toString();
			ControlName3="txtActualQty"+i.toString();
			ControlName4="txtComments"+i.toString();
			if(document.getElementById(ControlName1)!=null)
			{
				var Qty=parseFloat(document.getElementById(ControlName2).value);
				var ActualyQty=parseFloat(document.getElementById(ControlName3).value);
				var Comments=document.getElementById(ControlName4).value;
				if(Qty!=ActualyQty && Comments=="")
				{
					Flag=1;
					window.alert("实际数量和系统数量不符合,请输入备注信息!");
					return;
				}
			}
		}
		if(Flag==0)
		{
			form1.action="PrintLabel2.asp?Action=2";
			form1.submit();
		}
	 }
	 
	 function GoBack()
	 {
	 	window.parent.location="printlabel.asp"
	 }
</script>
</head>
<body bgcolor="#339966">
<table width="90%"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <form id="form1" method="post" name="form1">
    <tr>
      <td height="20" colspan="2" class="t-t-DarkBlue"  align="center">Label Print(入库标签打印)</td>
    </tr>
    
	 <tr>
      <td colspan="2">请核对子工单列表:<BR>
	  <%
	  	arrPrintSubJob=split(SubJob,",")
	  	for i=0 to ubound(arrPrintSubJob)-1
			response.write "<strong>" & arrPrintSubJob(i) &"</strong><br>"
		next 
		response.write "<strong>" & arrPrintSubJob(ubound(arrPrintSubJob)) &"</strong>"
		
		normalQty=GetQty(ManJobNumber,NewJob,"-1",normalQty)
		SlappingQty=GetQty(ManJobNumber,NewJob,"5",SlappingQty)
		ActualQty01=request("txtActualQty-1")
		ActualQty5=request("txtActualQty-5")
		if ActualQty01="" then
			ActualQty01=normalQty
		end if
		if ActualQty5="" then
			ActualQty5=SlappingQty
		end if
	   %>
	  
	  </td>
    </tr>
	<tr>
		<td>
		<table width="100%" colspan="2" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			<tr>
				<td>Type(类型)</td>
				<td>Qty(系统数量)</td>
				<td>Actualy Qty(实际数量)</td>
				<td>Comments(备注)</td>
			</tr>
				<tr>
				<td width="30%">Normal<input type="hidden" id="txtTransactionType-1" name="txtTransactionType-1" value="-1"></td>
				<td width="30%"><input type="hidden" id="txtQty-1" name="txtQty-1" value=<%=normalQty%>><%=normalQty%></td>
				<td width="30%"><input type="text" id="txtActualQty-1" name="txtActualQty-1" value=<%=ActualQty01%>></td>
				<td width="10%"><input type="text" id="txtComments-1" name="txtComments-1" value=<%=request("txtComments-1")%>></td>
			</tr>
			<%i=0
			 while not rs1.eof
			 	transQty=GetQty(ManJobNumber,NewJob,rs1("TRANSACTION_TYPE"),rs1("QTY"))   '得到不良品
				comments=request("txtComments"&rs1("TRANSACTION_TYPE"))
				actualQty=request("txtActualQty"&rs1("TRANSACTION_TYPE"))
				if actualQty = "" then
					actualQty=transQty
				end if
				transType="None"
				select case rs1("TRANSACTION_TYPE")
					case "0"
						transType="None"
					case "1"
						transType="Rework"					
					case "2"
						transType="Scrap"
					case "3"
						transType="Readjust"
					case "4"
						transType="Change Model"				
				end select
			 %>			
			<tr><td><%=transType%>         	 
			 <input type="hidden" id="txtTransactionType<%=rs1("TRANSACTION_TYPE")%>" name="txtTransactionType<%=rs1("TRANSACTION_TYPE")%>" value="<%=rs1("TRANSACTION_TYPE")%>">
			</td>
			<td><input type="hidden" id="txtQty<%=rs1("TRANSACTION_TYPE")%>" name="txtQty<%=rs1("TRANSACTION_TYPE")%>" value=<%=transQty%>><%=transQty%></td>
			<td><input type="text" id="txtActualQty<%=rs1("TRANSACTION_TYPE")%>" name="txtActualQty<%=rs1("TRANSACTION_TYPE")%>" value="<%=actualQty%>">
			</td>
			<td><input type="text" id="txtComments<%=rs1("TRANSACTION_TYPE")%>" name="txtComments<%=rs1("TRANSACTION_TYPE")%>" value="<%=comments%>"></td>
			</tr>
			<%	
				i=i+1
				rs1.movenext
			wend
			rs1.close
			set rs1=nothing
			%>
			<%IF SlappingQty>0 then%>
			<tr>
				<td width="30%">Slapping<input type="hidden" id="txtTransactionType5" name="txtTransactionType5" value="5"></td>
				<td width="30%"><input type="hidden" id="txtQty5" name="txtQty5" value=<%=SlappingQty%>><%=SlappingQty%></td>
				<td width="30%"><input type="text" id="txtActualQty5" name="txtActualQty5" value=<%=ActualQty5%>></td>
				<td width="10%"><input type="text" id="txtComments5"  name="txtComments5" value=<%=request("txtComments5")%>></td>
			</tr>
			<%end if%>
		</table>
		</td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr>
		<Td colspan="2" align="center">
			<input name="btnSubmit" type="button" id="btnSubmit"  value="Print(打印)" onClick="Print()">
			&nbsp;
			<input name="btnBack" type="button"  id="btnBack"  value="Back(返回)" onClick="GoBack()">
			&nbsp;
			<input name="btnClose" type="button"  id="btnClose"  value="Close(关闭)" onClick="window.close();">
		</Td>
	</tr>
	 <input type="hidden" id="txtSubJobList" name="txtSubJobList" value=<%=SubJob%>>
	 <input type="hidden" id="newJob" name="newJob" value=<%=replace(NewJob,",","-")%>>
	 <input type="hidden" id="batchNo" name="batchNo" value=<%=ManJobNumber%>>
	 <input type="hidden" id="TotalCount" name="TotalCount" value=<%=i%>>
	 <input type="hidden" id="txtOperatorCode" name="txtOperatorCode" value=<%=OperatorCode%>>
	   <input type="hidden" id="computername" name="computername" value=<%=ComputerName%>>
  </form>
</table>
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->