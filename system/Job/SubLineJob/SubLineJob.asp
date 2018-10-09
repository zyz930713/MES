<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Open.asp" -->
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Trans.asp" -->
<!--#include virtual="/Functions/RanPwd.asp" -->

<%	
	dim arry 
	dim isSearch
	isSearch="0"
	Action=request.QueryString("action")
	if (request.QueryString("action")="1") then
		Set Upload = Server.CreateObject("Persits.Upload")
		Upload.LogonUser "", "ASPEncrypt", "Dickens;'2006"
		Upload.CodePage = 936
		Upload.Save
	
		Set File = Upload.Files("file1")
		if not File is nothing then
			context_path="\JOB\SubLineJob\TempJobExcel"
			drive_path=server.mappath(context_path)
			org_file_name=File.FileName
			org_file_type=File.ContentType
			org_file_size=File.size
			file_name=gen_key(20)&gen_key(20)
			file_path=file_name&File.Ext
			
			if File.Ext<>".xls"  then
				response.Redirect("SubLineJob.asp?errorstring=请上传excel的文件!")
			end if
			
			File.SaveAs drive_path &"\"& file_path  '保存文件到服务器
		 
				 set excelconn=server.createobject("adodb.connection") 
				 'excelconn.open "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = "+drive_path &"\"& file_path+";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'" 
				 excelconn.open "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "+drive_path &"\"& file_path+";Extended Properties='Excel 8.0;HDR=NO;IMEX=1'" 

				 set rsExcel=server.createobject("adodb.recordset")
				 SQL="select * from [Sheet1$]" 
				 rsExcel.open SQL,excelconn,1,3
				arry= rsExcel.GetRows
				isSearch="1"
		else
			response.write "未能获得上传文件!"
		end if
	end if 
	
	if(request.QueryString("action")="2")then
		Set Upload2 = Server.CreateObject("Persits.Upload")
		Upload2.LogonUser "", "ASPEncrypt", "Dickens;'2006"
		Upload2.CodePage = 936
		Upload2.Save
		excelrecordcount=Upload2.Form("txtcount")

		importdatetime=now
		Set TypeLib = CreateObject("Scriptlet.TypeLib")
   		Guid = TypeLib.Guid
		 for mm=0 to excelrecordcount
				CompletionDate=Upload2.Form("txtcompletiondata"&cstr(mm))
				LineName=Upload2.Form("txtlinename"&cstr(mm))
				PartNumber=Upload2.Form("txtpartnumber"&cstr(mm))
				Qty=Upload2.Form("txtqty"&cstr(mm))
				JobType=Upload2.Form("txttype"&cstr(mm))
				JobPriority=Upload2.Form("txtpriority"&cstr(mm))
				jobnumber=GetSubJobNumber()
  
				set rsTBL_MES_LOTMASTER=server.createobject("adodb.recordset")
				SQL="insert into TBL_MES_LOTMASTER (ORGANIZATIONID,WIPENTITYID,WIPENTITYNAME,DESCRIPTION,primaryitemdesc,CLASSCODE,CREATEPERSON,DATECREATION,"
				SQL=SQL+" DATECOMPLETED,SCHSTARTDATE,SCHCOMPLETIONDATE,STARTQUANTITY,LINECODE,JOB_PRIORITY,JOB_TYPE)"
				SQL=SQL+"values('24','"+jobnumber+"','"+jobnumber+"','Sub Job','"+PartNumber+"','C10-WIP1','"+session("code")+"',to_date('"+cstr(importdatetime)+"','yyyy-mm-dd hh24:mi:ss'),"
				SQL=SQL+"to_date('"+CompletionDate+"','yyyy-mm-dd hh24:mi:ss'),to_date('"+CompletionDate+"','yyyy-mm-dd hh24:mi:ss'),to_date('"+CompletionDate+"','yyyy-mm-dd hh24:mi:ss'),"
				SQL=SQL+"'"+Qty+"','"+LineName+"','"+JobPriority+"','"+JobType+"')"
				rsTBL_MES_LOTMASTER.open SQL,connTicket,1,3
				
	 
				
				SQL="select * from ITEM_BOM WHERE PART_NUMBER='"+PartNumber+"'"
				set rsBOM=server.createobject("adodb.recordset")
				rsBOM.open SQL,conn,1,3
				j=0
				while not rsBOM.eof
					set rsTBL_LOT_BOM_ITEM=server.createobject("adodb.recordset")
					materialqty=cstr(cdbl(rsBOM("QTY"))*cdbl(Qty))
					SQL="INSERT INTO TBL_LOT_BOM_ITEM (INVENTORY_ITEM_ID, ORGANIZATION_ID, INVENTORY_ITEM_NAME, DESCRIPTION, WIP_ENTITY_ID, REQUIRE_QUANTITY, LAST_UPDATE_DATE,operation_seq_num)"
					SQL=SQL+"VALUES('"+jobnumber+"','24','"+rsBOM("SUB_PART_NUMBER")+"','"+rsBOM("DESCRIPTION")+"','"+jobnumber+"','"+cstr(materialqty)+"',sysdate,"+cstr(j)+")"
			 		rsTBL_LOT_BOM_ITEM.open SQL,connTicket,1,3
					rsBOM.movenext
					j=j+1
				wend 
				
				set rsSUB_JOB_LIST=server.createobject("adodb.recordset")
				SQL="SELECT * FROM SUB_JOB_LIST WHERE 1=2"
				rsSUB_JOB_LIST.open SQL,conn,1,3
				rsSUB_JOB_LIST.addnew
				rsSUB_JOB_LIST("GUID")=Guid
				rsSUB_JOB_LIST("JOB_NUMBER")=jobnumber
				rsSUB_JOB_LIST("SEQUENCE")=cint(cstr(cdbl(replace(jobnumber,"S",""))))
				rsSUB_JOB_LIST.update
			next 
			
			set rsSUB_JOB_IMPORT_HISTORY=server.createobject("adodb.recordset")
			SQL="SELECT * FROM SUB_JOB_IMPORT_HISTORY WHERE 1=2"

			rsSUB_JOB_IMPORT_HISTORY.open SQL,conn,1,3
			rsSUB_JOB_IMPORT_HISTORY.addnew
			rsSUB_JOB_IMPORT_HISTORY("GUID")=Guid
			rsSUB_JOB_IMPORT_HISTORY("IMPORT_DATETIME")=importdatetime
			rsSUB_JOB_IMPORT_HISTORY("IMPORT_USER_CODE")=session("code")
			rsSUB_JOB_IMPORT_HISTORY.update
			
			set rsSubJobNumber2=server.createobject("adodb.recordset")
			sql="select a.JOB_NUMBER from SUB_JOB_LIST a , SUB_JOB_IMPORT_HISTORY b where a.guid=b.guid and b.IMPORT_DATETIME=to_date('"+cstr(importdatetime)+"','yyyy-mm-dd hh24:mi:ss')"
			rsSubJobNumber2.open sql,conn,1,3
			jobnumberlist=""
			while not rsSubJobNumber2.eof
				jobnumberlist=jobnumberlist+"'"+rsSubJobNumber2(0)+"',"
				rsSubJobNumber2.movenext
			wend
			jobnumberlist=left(jobnumberlist,len(jobnumberlist)-1)
			
			SQL="select SCHCOMPLETIONDATE, linecode,primaryitemdesc, STARTQUANTITY,JOB_TYPE,JOB_PRIORITY, WIPENTITYNAME,DATECREATION AS  DATETIME "
			SQL=SQL+" from TBL_MES_LOTMASTER WHERE WIPENTITYNAME in ("+jobnumberlist+")"
			set rsJobNumber2=server.createobject("adodb.recordset")
			rsJobNumber2.open SQL,connTicket,1,3

			arry=rsJobNumber2.GetRows
			isSearch="2"
	 
	end if 
	
	
	if(request.QueryString("action")="3")then
			set rsSubJobNumber2=server.createobject("adodb.recordset")
			sql="select a.JOB_NUMBER from SUB_JOB_LIST a , SUB_JOB_IMPORT_HISTORY b where a.guid=b.guid"
			rsSubJobNumber2.open sql,conn,1,3
			jobnumberlist=""
			while not rsSubJobNumber2.eof
				jobnumberlist="'"+rsSubJobNumber2(0)+"',"
				rsSubJobNumber2.movenext
			wend
			jobnumberlist=left(jobnumberlist,len(jobnumberlist)-1)
			
			SQL="select SCHCOMPLETIONDATE, linecode,primaryitemdesc, STARTQUANTITY,JOB_TYPE,JOB_PRIORITY, WIPENTITYNAME,DATECREATION AS  DATETIME "
			SQL=SQL+" from TBL_MES_LOTMASTER WHERE 1=1"
			SQL=SQL+" AND WIPENTITYNAME like 'S%' ORDER BY WIPENTITYNAME"
			set rsJobNumber2=server.createobject("adodb.recordset")
			rsJobNumber2.open SQL,connTicket,1,3
		
			arry=rsJobNumber2.GetRows
			isSearch="3"
	end if 
	

%>

<%
	function GetSubJobNumber()
		set rsSubJobNumber=server.createobject("adodb.recordset")
		rsSubJobNumber.open "select decode(max(SEQUENCE),null,0,max(SEQUENCE)) from SUB_JOB_LIST",conn,1,3
		
		if(rsSubJobNumber(0)="0") then
			GetSubJobNumber="S0000001"
		else
			SubJobNumber=cstr(rsSubJobNumber(0))
		
			SubJobNumber=cstr(cdbl(SubJobNumber)+1)
			
			for i=0 to 6-len(cstr(SubJobNumber))
				SubJobNumber="0"+cstr(SubJobNumber)
			next 
			GetSubJobNumber="S"+SubJobNumber
		end if 
	end function 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Scan </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script>
	function importExcel()
	{
		form1.action="SubLineJob.asp?Action=1";
		form1.submit();
	}
	
	function showMaterial(Material_Part_Number,JobNumber,Qty)
	{
		window.showModalDialog("Material_BOM.asp?Material_Part_Number="+Material_Part_Number+"&JobNumber="+JobNumber+"&qty="+Qty,window,"dialogHeight:800px;dialogWidth:800px");
		//window.open("Material_BOM.asp?Material_Part_Number="+Material_Part_Number+"&JobNumber="+JobNumber+"&qty="+Qty);
		
	}
	function strDateTime(str) 
	{ 
		var reg = /^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/; 
		var r = str.match(reg); 
		if(r==null)return false; 
		var d= new Date(r[1], r[3]-1,r[4],r[5],r[6],r[7]); 
		return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]&&d.getHours()==r[5]&&d.getMinutes()==r[6]&&d.getSeconds()==r[7]); 
	}

	function SaveExcel()
	{
		var count=document.getElementById("txtcount").value;
		for(var i=0; i<=count;i++)
		{
			var importdatetime=document.getElementById("txtcompletiondata"+i).value
			if(!strDateTime(importdatetime))
			{
				alert((i+1)+"行的日期格式不对！");
				return;
			}
			
			if (isNaN(document.getElementById("txtqty"+i).value))
			{
				alert((i+1)+"行的数量应该为整型！");
				return;
			}
			
			if (document.getElementById("txttype"+i).value=="")
			{
				alert("请选择"+(i+1)+"行的Job Type!");
				return;
			}
			
			if (document.getElementById("txtpriority"+i).value=="")
			{
				alert("请选择"+(i+1)+"行的Job Priority!");
				return;
			}

		}
		form1.action="SubLineJob.asp?Action=2";
		form1.submit();
	}
	
	function Search()
	{
		form1.action="SubLineJob.asp?Action=3";
		form1.submit();
	}
</script>
</head>

<body>
<form  name="form1"  id="form1" method="post" enctype="multipart/form-data">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Sub Line Job</td>
  </tr>
  <tr>
    <td height="20">Job Number </td>
     <td height="20">
	 	 <input type="text" name="txtjobnumber" id="txtjobnumber" value="<%=jobnumber%>">
	</td>
    <td>Part Number</td>
     <td height="20"> 
	 	<input type="text" name="txtmodelname" id="txtmodelname" value="<%=modelname%>">
	 </td>
  
    <td>Line Name</td>
    <td height="20"> 
	 	<input type="text" name="txtlinename" id="txtlinename" value="<%=linename%>">
	 </td>
	</tr>
	<tr>
	 <td>Import DateTime From:</td> 
	 <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">
	
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
	<select name="fromtime" id="fromtime">
   			 <option value="14:30:00" <% if fromtime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if fromtime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
</td>
<td>Import DateTime To:</td>
<td>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="16">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
	<select name="totime" id="totime">
   			 <option value="14:30:00" <% if totime="14:30:30" then response.write "Selected" end if%>>14:30:00</option>
			  <option value="23:59:59" <% if totime="23:59:59" then response.write "Selected" end if%>>23:59:59</option>
	   			 
  		</select>  
	
&nbsp;
	<Td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="Search()"></Td>
  	<Td>&nbsp;</Td>
  </tr>

	<tr>
  	<Td class="t-t-Borrow"  colspan="8">Select Excel file to import</Td>
  </tr>
  
  <tr>
  	<Td>File Name</Td>
	<Td><input name="file1" type="file" id="file1" size="100"></Td>
  	<Td><input type="button" id ="import" name="import" value="Import Excel" onclick="importExcel()" class="t-b-Yellow" style="width:100px"></Td>
	<Td colspan="5">
		<%if Action<>"3" then%>
		<input type="button" id ="import" name="import" value="Save" onclick="SaveExcel()" class="t-b-Yellow" style="width:100px">
		<%end if %>
		</Td>
  </tr>
</table>

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
  	 <td height="20" class="t-t-Borrow">Completion Date</td>
	  <td height="20"class="t-t-Borrow">Line Name</td>
	 <td height="20"class="t-t-Borrow">Part Number</td>
	  <td height="20" class="t-t-Borrow">Qty</td>
	  <td height="20" class="t-t-Borrow">Job Type</td>
	  <td height="20" class="t-t-Borrow">Job Priority</td>
	   <td height="20" class="t-t-Borrow">Job Number</td>
	  <td height="20" class="t-t-Borrow">Create Time</td>
	   <td height="20" class="t-t-Borrow">Material BOM</td>
  </tr>
  
  <%
  		function getJobPriority(priorityid)
				set rsJobPriority=server.createobject("adodb.recordset")
				sql="select  * from job_priority_setting where PRIORITY_LEVEL='"+cstr(priorityid)+"'"
				rsJobPriority.open sql,conn,1,3
				if(rsJobPriority.recordcount>0 ) then
					getJobPriority=rsJobPriority(1)
				else
					getJobPriority=""
				end if 
		end function 
  %>
  <%if isSearch="1" then
  	For row = 0 To UBound(arry,2)
  %>
  <tr>
  	 <td height="20"><input id="txtcompletiondata<%=row%>" name="txtcompletiondata<%=row%>" type="input" value="<%=arry(0, row)%>"></td>
	  <td height="20"><input id="txtlinename<%=row%>" name="txtlinename<%=row%>" type="input" value="<%=arry(1, row)%>"></td>
	 <td height="20"><input id="txtpartnumber<%=row%>" name="txtpartnumber<%=row%>" type="input" value="<%=arry(2, row)%>"></td>
	  <td height="20"><input id="txtqty<%=row%>" name="txtqty<%=row%>" type="input" value="<%=arry(3, row)%>"></td>
	
	  <td height="20" >
			<%
				SQL="SELECT * FROM JOB_TYPE_SETTING order by JOB_TYPE"
				set rsJobType=Server.CreateObject("adodb.recordset")
				rsJobType.open SQL,conn,1,3
			%>
				<select name="txttype<%=row%>"  id="txttype<%=row%>"tabindex="0" style="width:100px">
					  <option value="">--Select--</option>
					  <% while not rsJobType.eof%>
						<option value=<%=rsJobType("JOB_TYPE")%> <%if trim(arry(4, row))=rsJobType("JOB_TYPE") then response.write " selected " end if %>><%=rsJobType("JOB_TYPE")%></option>
					  <%
							rsJobType.movenext
						wend
					  %>
			</select>
	  </td>

	   <td height="20" >
	   	<%
			SQL="SELECT * FROM JOB_PRIORITY_SETTING order by PRIORITY_LEVEL"
			set rsPriority=Server.CreateObject("adodb.recordset")
			rsPriority.open SQL,conn,1,3
		%>
	   		<select name="txtpriority<%=row%>"  id="txtpriority<%=row%>"tabindex="0" style="width:100px">
				  <option value="">--Select--</option>
				  <% while not rsPriority.eof%>
				  	<option value=<%=rsPriority("PRIORITY_LEVEL")%> <%if trim(arry(5, row))=trim(rsPriority("PRIORITY_LEVEL")) then response.write " selected " end if %>><%=rsPriority("PRIORITY_DEC")%></option>
				  <%
				  		rsPriority.movenext
				  	wend
				  %>
		</select>
	   </td>
	   
	   <td height="20">&nbsp;</td>
	  <td height="20"><%=now%></td>
	   <td height="20"><input id="btnBOM<%=row%>" name="btnBOM<%=row%>" type="button" value="BOM" onclick="showMaterial('<%=arry(2, row)%>','0','<%=arry(3, row)%>')"></td>
  </tr>		 
	<%			
		Next
	 end if %>
	 
	<%if isSearch="2" or isSearch="3" then
  	For row = 0 To UBound(arry,2)
  %>
  <tr>
  	 <td height="20"><%=arry(0, row)%></td>
	 <td height="20"><%=arry(1, row)%></td>
	 <td height="20"><%=arry(2, row)%></td>
	 <td height="20"><%=arry(3, row)%></td>
	 <td height="20"><%=arry(4, row)%></td>
	 <td height="20"><%=getJobPriority(cstr(arry(5, row)))%></td>
	 <td height="20"><%=arry(6, row)%></td>
	 <td height="20"><%=arry(7, row)%></td>
	  <td height="20"><input id="btnBOM<%=row%>" name="btnBOM<%=row%>" type="button" value="BOM" onclick="showMaterial('<%=arry(2, row)%>','<%=arry(6, row)%>','<%=arry(3, row)%>')"></td>
  </tr>		 
	<%			
		Next
	 end if %>
	 
</table>
 <%if isSearch="1" then%>
<input id="txtcount" name="txtcount" type="hidden" value="<%=UBound(arry,2)%>">
<%else%>
<input id="txtcount" name="txtcount" type="hidden" value="0">
<%end if %>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Commit.asp" -->
<!--#include virtual="/WOCF/BOCF_TICKET_Close.asp" -->
<!--#include virtual="/WOCF/BOCF_Close.asp" -->