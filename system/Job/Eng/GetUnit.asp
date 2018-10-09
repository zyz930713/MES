<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include virtual="/Components/CacheControl.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
	JobNumber=trim(request("txtJobNumber"))
	SheetNumber=trim(request("txtSheetNumber"))
	TransactionType=request("selType")
	Qty=request("txtQty")
	UserCode=session("code")
	
	UserCodeSearch=request("txtUserCodeSearch")
	JobNumberSearch=request("txtJobNumberSearch")
	SheetNumberSearch=request("txtSheetNumberSearch")
	TransactionTypeSearch=request("selTypeSearch")
	fromdate=request("fromdate")
	todate=request("todate")
	set rsSearch=server.CreateObject("adodb.recordset")
	
	if(request.querystring("Action")="1") then
 		
		set rsIsInStore=server.CreateObject("adodb.recordset")
 		SQL="select * from offline_store where job_number='"+ JobNumber +"' and subjoblist like '%"+CSTR(CINT(SheetNumber))+"%' and TRANSACTIONTYPE='"+TransactionType+"' and isprint='1'"
		
		rsIsInStore.open SQL,conn,1,3
		if(rsIsInStore.recordcount>0) then
			response.write("<script>window.alert('此工单已经开始修理，不能借!')</script>")
		else
				'add qty check
			SQL=" SELECT decode(SUM(JD.DEFECT_QUANTITY),null,0,SUM(JD.DEFECT_QUANTITY)) AS QTY FROM JOB_DEFECTCODES JD , DEFECTCODE DC WHERE JD.DEFECT_CODE_ID=DC.NID(+)"
			SQL=SQL+" AND JD.JOB_NUMBER='"+TRIM(JobNumber)+"' AND JD.SHEET_NUMBER='"+CSTR(CINT(SheetNumber))+"' AND DC.TRANSACTION_TYPE =1"
			set rsCheckQty=server.CreateObject("adodb.recordset")
			rsCheckQty.open sql,conn,1,3
			
			SQL=" SELECT decode(SUM(UNITQTY),null,0,SUM(UNITQTY)) AS QTY FROM ENG_BORROW  WHERE JOBNUMBER='"+TRIM(JobNumber)+"' AND SHEETNUMBER='"+CSTR(CINT(SheetNumber))+"' AND TRANSACTIONTYPE =1"
			set rsCheckQty2=server.CreateObject("adodb.recordset")
			rsCheckQty2.open sql,conn,1,3
			
		
			if(rsCheckQty.recordcount>0) then
				if cint(cstr(rsCheckQty(0)))<cint(cstr(Qty)+cint(cstr(rsCheckQty2(0)))) then
					response.write("<script>window.alert('此工单的Rework数量为"+cstr(rsCheckQty(0))+",已经借出"+cstr(rsCheckQty2(0))+",不能借"+cstr(Qty)+"!')</script>")
				 else
					Set TypeLib = CreateObject("Scriptlet.TypeLib")
					ID = TypeLib.Guid
					SQL="select * from ENG_BORROW where 1<>1"
					rs.open SQL,conn,1,3
					if rs.recordcount=0 then
						rs.addnew
						rs("ID")=ID
						rs("USERCODE")=UserCode
						rs("BORROWTIME")=NOW()
						rs("UNITQTY")=Qty
						rs("JOBNUMBER")=TRIM(JobNumber)
						rs("SHEETNUMBER")=cint(SheetNumber)
						rs("TRANSACTIONTYPE")=TransactionType
						rs("STATION_ID")=""								
						rs.update
								 
						response.write("<script>window.alert('Borrow Unit Successfully!')</script>")
					end if
								rs.close
							'update by jack zhang 2011-1-7
							'SQL="UPDATE offline_store SET QTY=QTY-"&Qty&",ACTUALQTY=ACTUALQTY-"&Qty&" WHERE job_number='"+ JobNumber +"' and subjoblist like '%"+CSTR(CINT(SheetNumber))+"%' and TRANSACTIONTYPE='"+TransactionType+"'"
							SQL="UPDATE offline_store SET QTY=QTY-"&Qty&",ACTUALQTY=ACTUALQTY-"&Qty&" WHERE job_number='"+ JobNumber +"'  and TRANSACTIONTYPE='"+TransactionType+"'"
							SQL=SQL+"AND (subjoblist like '%-"+CSTR(CINT(SheetNumber))+"-%' or subjoblist like '"+CSTR(CINT(SheetNumber))+"-%' or subjoblist like '%-"+CSTR(CINT(SheetNumber))+"' or subjoblist = '"+CSTR(CINT(SheetNumber))+"') "

							set rsOfflineQty=server.CreateObject("adodb.recordset")
							rsOfflineQty.open SQL,conn,1,3
						end if 
					end if 
			end if
	end if	
	
	
	if(request.querystring("Action")="2") then
		SQL="select USERCODE,JOBNUMBER,SHEETNUMBER,TRANSACTIONTYPE,BORROWTIME,UNITQTY from ENG_BORROW where 1=1"
		if(UserCodeSearch<>"") then
			SQL=SQL+" AND USERCODE='"+UserCodeSearch+"'"
		end if 
		if(JobNumberSearch<>"") then
			SQL=SQL+" AND JOBNUMBER='"+JobNumberSearch+"'"
		end if 
		if(SheetNumberSearch<>"") then
			SQL=SQL+" AND SHEETNUMBER='"+SheetNumberSearch+"'"
		end if 
		if(TransactionTypeSearch<>"" and TransactionTypeSearch<>"-2") then
			SQL=SQL+" AND TRANSACTIONTYPE='"+TransactionTypeSearch+"'"
		end if 
			
		if(fromdate<>"") then
			SQL=SQL+" AND BORROWTIME>=to_date('yyyy-mm-dd hh24:mi:ss','"+fromdate+" 00:00:00')"
		end if 
		
		if(todate<>"") then
			SQL=SQL+" AND BORROWTIME<=to_date('yyyy-mm-dd hh24:mi:ss','"+todate+" 23:59:59')"
		end if 
	
		rsSearch.open SQL,conn,1,3
		
	end if	
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System - Engineer Borrow </title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
<script language="javascript" type="text/javascript">
 	function SaveData()
	{
		 if(document.getElementById("txtJobNumber").value==""  &&  document.getElementById("txtSheetNumber").value==""  && document.getElementById("txtQty").value=="")
		{
			window.alert("Please key in Job Number and Sheet Number");
			return;
		}	 	
		 /*if(document.getElementById("txtJobNumber").value!="" && document.getElementById("txtQty").value!="")
		{
			window.alert("Please key in Job Number or Batch Number");
			return;
		}	 */	
		 
		form1.action="GetUnit.asp?Action=1";
		form1.submit();
	}
	
	function SearchData()
	{
		form1.action="GetUnit.asp?Action=2";
		form1.submit();
	}
</script>

</head>
<body>
<form method="post" name="form1"  id="form1">
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="8" class="t-b-midautumn">Engineer Sample</td>
  </tr>
  <tr>
  	 <td>JobNumber</td>
	 <td><input name="txtJobNumber" type="text" id="txtJobNumber" value="<%=JobNumber%>" size="16"></td>
	 <td>SheetNumber</td>
	 <td><input name="txtSheetNumber" type="text" id="txtSheetNumber" value="<%=SheetNumber%>" size="16"></td>
  	 <td>Transaction Type</td>
	 <td>
	 	<select id="selType"  name="selType" style="width:100px">
			<option value="1" <%if TransactionType="1" then%> selected <% end if%> >Rework</option>
			<option value="4" <%if TransactionType="4" then%> selected <% end if%>>Change Model</option>
		</select>
	</td>
	 <td>Qty</td>
	 <td><input name="txtQty" type="text" id="txtQty" value="<%=Qty%>" size="16"></td>
  </tr>
  <tr>
		<Td colspan="8" align="center"><input name="btnSave" type="button" class="t-b-Yellow" id="btnSave" onClick="SaveData()" value="Save (保存)">
		</Td>
  </tr>
  <tr>
  	<Td colspan="8">
		History Info
	</Td>
  </tr>
   <tr>
  	<Td>User Code</Td>
	<td><input name="txtUserCodeSearch" type="text" id="txtUserCodeSearch" value="<%=UserCodeSearch%>" size="16"></td>
	<td>JobNumber</td>
	<td><input name="txtJobNumberSearch" type="text" id="txtJobNumberSearch" value="<%=JobNumberSearch%>" size="16"></td>
	<td>SheetNumber</td>
	<td><input name="txtSheetNumberSearch" type="text" id="txtSheetNumberSearch" value="<%=SheetNumberSearch%>" size="16"></td>
  	<td>Transaction Type</td>
	<td>
	 	<select id="selTypeSearch"  name="selTypeSearch" style="width:100px">
			<option value="-2" <%if TransactionTypeSearch="-2" then%> selected <% end if%>>--All--</option>
			<option value="-1" <%if TransactionTypeSearch="-1" then%> selected <% end if%>>Normal</option>
			<option value="0" <%if TransactionTypeSearch="0" then%> selected <% end if%>>None</option>
			<option value="1" <%if TransactionTypeSearch="1" then%> selected <% end if%> >Rework</option>
			<option value="2" <%if TransactionTypeSearch="2" then%> selected <% end if%>>Scrap</option>
			<option value="3" <%if TransactionTypeSearch="3" then%> selected <% end if%>>Readjust</option>
			<option value="4" <%if TransactionTypeSearch="4" then%> selected <% end if%>>Others</option>
			<option value="5" <%if TransactionTypeSearch="5" then%> selected <% end if%>>Slapping</option>
		</select>
		
	</td>
  </tr>
  <tr>
  	 <td>DateTime From:</td> 
	 <td><input name="fromdate" type="text" id="fromdate" value="<%=fromdate%>" size="16">
	
        <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
</td>
<td>To:</td>
<td>
<input name="todate" type="text" id="todate" value="<%=todate%>" size="16">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
	 
	
&nbsp;
</td>
<Td colspan="4">
	<input name="btnSearch" type="button" class="t-b-Yellow" id="btnSearch" onClick="SearchData();" value="Search(查询)">
</Td>
</tr>
<tr>
	<Td colspan="8">
		<Table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
			 <TR>
			 	<TD class="t-t-Borrow">USERCODE</TD>
				<TD class="t-t-Borrow">JOB_NUMBER</TD>
				<TD class="t-t-Borrow">SHEET_NUMBER</TD>
				<TD class="t-t-Borrow">TRANSACTIONTYPE</TD>
				<TD class="t-t-Borrow">TIME</TD>
				<TD class="t-t-Borrow">QTY</TD>
			</TR>
			
			<%
				IF request.querystring("Action")="2" then
				while not rsSearch.eof
				
			%> 
				<Tr>
				<TD><%=rsSearch("USERCODE")%></TD>
				<TD><%=rsSearch("JOBNUMBER")%></TD>
				<TD><%=rsSearch("SHEETNUMBER")%></TD><br>
				<TD>
					<% IF rsSearch("TRANSACTIONTYPE")="-1" THEN RESPONSE.WRITE "Normal" END IF%>	
					<% IF rsSearch("TRANSACTIONTYPE")="0" THEN RESPONSE.WRITE "None" END IF%>	
					<% IF rsSearch("TRANSACTIONTYPE")="1" THEN RESPONSE.WRITE "Rework" END IF%>	
					<% IF rsSearch("TRANSACTIONTYPE")="2" THEN RESPONSE.WRITE "Scrap" END IF%>	
					<% IF rsSearch("TRANSACTIONTYPE")="3" THEN RESPONSE.WRITE "Readjust" END IF%>	
					<% IF rsSearch("TRANSACTIONTYPE")="4" THEN RESPONSE.WRITE "Others" END IF%>	
					<% IF rsSearch("TRANSACTIONTYPE")="5" THEN RESPONSE.WRITE "Slapping" END IF%>	
				</TD>
				<TD><%=rsSearch("BORROWTIME")%></TD>
				<TD><%=rsSearch("UNITQTY")%></TD>
				</Tr>
			<% 	rsSearch.movenext
				wend
				end if %>
			
		</Table>
	</Td>
</tr>
</table>
</form>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->