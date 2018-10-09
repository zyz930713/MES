﻿<%
response.charset = "UTF-8"
'if session("ShiftEditUser") = "" then call sussLoctionHref("网络超时或者您还没有登录请登录","user_login.asp")
'if session("ShiftEditType") <> "HV" then response.redirect "default.asp"
'ReadonlyStr = "readonly"
'if session("ShiftEditRole") = "0" then ReadonlyStr = ""
%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual= "/Functions/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>HV - Editor - Line</title>
<link rel="stylesheet" type="text/css" href="/css/basic.css">
<script language="javascript" type="text/javascript" src="/Scripts/ThreeMenu.js"></script>
<script language="javascript" type="text/javascript" src="/Scripts/DatePicker/WdatePicker.js"></script>
<script type="text/javascript">
function $(obj){
   return document.getElementById(obj);
}
function beforeSubmit(form)
{	if(document.InputHv.JobNumber.value.length==0){
     alert("工单不能为空!");
     document.InputHv.JobNumber.focus();
     return false;
    }
	if(document.InputHv.Adate.value.length==0){
     alert("时间不能为空!");
     document.InputHv.Adate.focus();
     return false;
    }
	
	if(document.InputHv.DName.value.length==0){
     alert("请选择白夜班!");
     document.InputHv.DName.focus();
     return false;
    }
return true;
}

function jsdt(){
	var OpenTime = parseFloat($('OpenTime').value);
	var StopTime = parseFloat($('StopTime').value);
	t1 = 720 - OpenTime;
	t1 = t1.toFixed(0);
	$('StopTime').value=t1;
}
function jst1f(){
	var t1g = parseFloat($('T1G').value);
	var t1b = parseFloat($('T1B').value);
	t1 = t1b + t1g;
	t1 = t1b / t1;
	t1 = t1 * 100;
	t1 = t1.toFixed(2);
	$('T1F').value=t1;
}
function jscf(cellno){
	var cg = 'C' + cellno + 'G';
	var cb = 'C' + cellno + 'B';
	var cf = 'C' + cellno + 'F';
	var t1g = parseFloat($(cg).value);
	var t1b = parseFloat($(cb).value);
	t1 = t1b + t1g;
	t1 = t1b / t1;
	t1 = t1 * 100;
	t1 = t1.toFixed(2);
	$(cf).value=t1;
}
</script>
</head>
<style type="text/css">
table tr{
	height:26px;
	line-height:26px;
	background:#F2F2F2;
}
table tr td{
	text-align:center;
}
</style>
<body style="margin:20px;font-size:14px;background-color:#339966;">

<%
	ProductName = request("ProductName")
	
	LineName = request("LineName")
	
	JobNumber=request("JobNumber")
	
	ShiftName = request("ShiftName")
	SearchEdit = request("SearchEdit")
	Adate = request("Adate")
	DName = request("DName")
	CellAddModi = request("CellAddModi")
	
	if Adate = "" then
		Adate = FormatTime(date(),2)
	end if
	
	'添加修改各Cell数据
	if CellAddModi = "添加" or CellAddModi = "修改" then
		JobNumber = Ucase(request("JobNumber"))
		
		if JobNumber = "" then call errorHistoryBack("工单号码不能为空")
		
		if instr(JobNumber,"-")=0 then
						errorHistoryBack("工单错误，请扫描子工单！")
		end if
		
		ajobnumber=split(JobNumber,"-")
		strSheetNo=ajobnumber(ubound(ajobnumber))

		if len(strSheetNo)>3 then
			errorHistoryBack("细分工单单编号超过3位错误，请扫描子工单！")
		end if
		
	if isnumeric(strSheetNo) then
	sheetnumber = cint(strSheetNo)
	else
	errorHistoryBack("工单错误，请重新扫描子工单！")
	end if
	job_number=left(JobNumber,len(JobNumber)-4)
		
	
		
		OpenTime = request("OpenTime")
		StopTime = request("StopTime")

		T1G = request("T1G")
		T1B = request("T1B")
		
		C1G = request("C1G")
		C1B = request("C1B")
		
		C2G = request("C2G")
		C2B = request("C2B")
		
		C3G = request("C3G")
		C3B = request("C3B")
		
		C4G = request("C4G")
		C4B = request("C4B")
		
		C5G = request("C5G")
		C5B = request("C5B")
		
		PrBlocked = request("PrBlocked")
		JOB_START_QUANTITY= clng(T1G)+clng(T1B)	
		if CellAddModi = "添加" then
	sql="select * from  report_hv_data WHERE Adate = '"&cdate(Adate)&"' AND ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' AND DName = '"&DName&"'"
		
		rs.open sql,conn,1,3
		if rs.bof then
			CellInsertSql = "INSERT INTO report_hv_data(adate,udate,productname,linename,dname,ShiftName,JOB_NUMBER,opentime,stoptime,t1G,t1B,c1G,c1B,c2G,c2B,c3G,c3B,c4G,c4B,c5G,c5B,PrBlocked)"
			CellInsertSql = CellInsertSql & " VALUES('"&cdate(Adate)&"',sysdate,'"&ProductName&"','"&LineName&"','"&DName&"','"&ShiftName&"','"&JobNumber&"','"&OpenTime&"','"&StopTime&"','"&T1G&"','"&T1B&"','"&C1G&"','"&C1B&"','"&C2G&"','"&C2B&"','"&C3G&"','"&C3B&"','"&C4G&"','"&C4B&"','"&C5G&"','"&C5B&"','"&PrBlocked&"')"
			'response.write CellInsertSql
			'response.end
			conn.execute(CellInsertSql)
			
			call UpdateJob
	    else
		CellUpdateSql = "UPDATE report_hv_data set Udate = sysdate,ShiftName = '"&ShiftName&"',OpenTime = '"&OpenTime&"',StopTime = '"&StopTime&"',JOB_NUMBER = '"&JobNumber&"'"
			CellUpdateSql = CellUpdateSql & ",T1G = '"&T1G&"',T1B = '"&T1B&"',C1G = '"&C1G&"',C1B = '"&C1B&"',C2G = '"&C2G&"',C2B = '"&C2B&"',C3G = '"&C3G&"',C3B = '"&C3B&"',C4G = '"&C4G&"',C4B = '"&C4B&"',C5G = '"&C5G&"',C5B = '"&C5B&"',PrBlocked = '"&PrBlocked&"'"
			CellUpdateSql = CellUpdateSql & " WHERE Adate = '"&cdate(Adate)&"' AND ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' AND DName = '"&DName&"'"
			'response.write CellUpdateSql
			'response.end
			conn.execute(CellUpdateSql)
			
			call UpdateJob
		end if	
		rs.close	
		response.Write("<script type='text/javascript'>")
		response.Write("window.close();")
        response.Write("</script>)")
			
			'call sussLoctionHref("添加成功","InputHv.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName&"&ShiftName="&ShiftName&"&JobNumber="&JobNumber)
		elseif CellAddModi = "修改" then
			CellUpdateSql = "UPDATE report_hv_data set Udate = sysdate,ShiftName = '"&ShiftName&"',OpenTime = '"&OpenTime&"',StopTime = '"&StopTime&"',JOB_NUMBER = '"&JobNumber&"'"
			CellUpdateSql = CellUpdateSql & ",T1G = '"&T1G&"',T1B = '"&T1B&"',C1G = '"&C1G&"',C1B = '"&C1B&"',C2G = '"&C2G&"',C2B = '"&C2B&"',C3G = '"&C3G&"',C3B = '"&C3B&"',C4G = '"&C4G&"',C4B = '"&C4B&"',C5G = '"&C5G&"',C5B = '"&C5B&"',PrBlocked = '"&PrBlocked&"'"
			CellUpdateSql = CellUpdateSql & " WHERE Adate = '"&cdate(Adate)&"' AND ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' AND DName = '"&DName&"'"
			'response.write CellUpdateSql
			response.end
			conn.execute(CellUpdateSql)
			
			call UpdateJob
			
			call sussLoctionHref("修改成功","InputHv.asp?SearchEdit=ok&ProductName="&ProductName&"&LineName="&LineName&"&Adate="&Adate&"&DName="&DName&"&ShiftName="&ShiftName)
		end if
	end if
	'添加修改各Cell数据
%>
<center>
<form name="InputHv" method="post" action="InputHV.asp" >
<table width="1000" border="1" cellspacing="0" cellpadding="0"><tr><td align="center" class="TitleLable"><p style="height:34px;line-height:34px;"><strong>线上设备产出</strong></p></td></tr>
	
</table>
</br>
<%
	'if SearchEdit = "刷新" or SearchEdit = "ok" then

		if IsNumeric(LineName) then
			set CTRS = Server.CreateObject("adodb.recordset")
			CTSQL = "select T1,C1,C2,C3,NVL(C4,0) C4,NVL(C5,0) C5 from report_hv_target"
			CTSQL = CTSQL & " WHERE ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' and Tdate < = '"&Adate&"' ORDER BY Tdate desc"
			'response.write CTSQL
			'response.end
			CTRS.open CTSQL,conn,1,1
			if not CTRS.eof then
				T1T = Formatnumber(CTRS("T1"),2,-1)
				C1T = Formatnumber(CTRS("C1"),2,-1)
				C2T = Formatnumber(CTRS("C2"),2,-1)
				C3T = Formatnumber(CTRS("C3"),2,-1)
				C4T = Formatnumber(CTRS("C4"),2,-1)
				C5T = Formatnumber(CTRS("C5"),2,-1)
			end if
			CTRS.CLOSE
			set CTRS = nothing

			set cellrs = Server.CreateObject("adodb.recordset")
			CellSql = "select productname,linename,dname,shiftname,JOB_NUMBER,nvl(opentime,0) opentime,nvl(stoptime,0) stoptime,nvl(t1G,0) T1G,nvl(t1B,0) T1B,nvl(c1G,0) c1g,nvl(c1B,0) c1b,nvl(c2G,0) c2g,nvl(c2B,0) c2b,nvl(c3G,0) c3g,nvl(c3B,0) c3b,nvl(c4G,0) c4g,nvl(c4B,0) c4b,nvl(c5G,0) c5g,nvl(c5B,0)C5B,nvl(PrBlocked,0) PrBlocked from report_hv_data"
			CellSql = CellSql & " WHERE Adate = '"&Adate&"' AND ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' AND DName = '"&DName&"'"
			'response.write CellSql
			'response.end
			cellrs.open CellSql,conn,1,1
				OpenTime = 720
				StopTime = 0
				PrBlocked = 0
				T1G = 0
				T1B = 0
				T1F = 0
				C1G = 0
				C1B = 0
				C1F = 0
				C2G = 0
				C2B = 0
				C2F = 0
				C3G = 0
				C3B = 0
				C3F = 0
				C4G = 0
				C4B = 0
				C4F = 0
				C5G = 0
				C5B = 0
				C5F = 0
				CellSubmit = "<input id = 'CellAddModi' type='submit' name='CellAddModi' value='添加' style='background:#FFBFBF;' onClick='return beforeSubmit()'/>"
			if not cellrs.eof then
				CellSubmit = "<input id = 'CellAddModi' type='submit' name='CellAddModi' value='修改' style='background:#BFFFBF;' onClick='return beforeSubmit()'/>"
				ShiftName1 = cellrs("ShiftName")
				JobNumber = cellrs("JOB_NUMBER")
				OpenTime = cellrs("OpenTime")
				StopTime = cellrs("StopTime")
				PrBlocked = cellrs("PrBlocked")
				T1G = CSNG(cellrs("T1G"))
				T1B = CSNG(cellrs("T1B"))
				T1A = T1G + T1B
				if T1A = 0 THEN
					T1F = 0
				ELSE
					T1F = Formatnumber(T1B / T1A * 100,2,-1)
				end if
				
				C1G = CSNG(cellrs("C1G"))
				C1B = CSNG(cellrs("C1B"))
				C1A = C1G + C1B
				if C1A = 0 THEN
					C1F = 0
				ELSE
					C1F = Formatnumber(C1B / C1A * 100,2,-1)
				END IF
				
				C2G = CSNG(cellrs("C2G"))
				C2B = CSNG(cellrs("C2B"))
				C2A = C2G + C2B
				if C2A = 0 THEN
					C2F = 0
				ELSE
					C2F = Formatnumber(C2B / C2A * 100,2,-1)
				END IF
				
				C3G = CSNG(cellrs("C3G"))
				C3B = CSNG(cellrs("C3B"))
				C3A = C3G + C3B
				if C3A = 0 THEN
					C3F = 0
				ELSE
					C3F = Formatnumber(C3B / C3A * 100,2,-1)
				END IF
				
				C4G = CSNG(cellrs("C4G"))
				C4B = CSNG(cellrs("C4B"))
				C4A = C4G + C2B
				if C4A = 0 THEN
					C4F = 0
				ELSE
					C4F = Formatnumber(C4B / C4A * 100,2,-1)
				END IF
				
				C5G = CSNG(cellrs("C5G"))
				C5B = CSNG(cellrs("C5B"))
				C5A = C5G + C5B
				if C5A = 0 THEN
					C5F = 0
				ELSE
					C5F = Formatnumber(C5B / C5A * 100,2,-1)
				END IF
			
			end if
			cellrs.close
			set cellrs = nothing
		end if
	'response.write ShiftName1
%>
			<table width="1000" border="1" cellspacing="0" cellpadding="0">
			<tr>
				<td colspan="5" align="center">
				<strong>
				产品：<font color="red">
				<input type="text" name="ProductName" id="ProductName" value="<%=ProductName%>" size="3" readonly/>
				</font>　线别:<input type="text" name="LineName" id="LineName" value="<%=LineName%>" size="3" readonly/>　日期：<input type="text" name="Adate" class="Wdate" id="Adate" onFocus="WdatePicker({isShowWeek:true,onpicked:function() {$dp.$('Adate_1').value=$dp.cal.getP('W','W');$dp.$('Adate_2').value=$dp.cal.getP('W','WW');}})" value="" size="10" />　D/N：<input name="DName" type="text" size="10" >
				<select name="DNameA"  onChange="(document.InputHv.DName.value=this.options[this.selectedIndex].value)">
				<option value="请选择白夜班">请选择白夜班</option>				  
				<option value="D" >Day</option>
				<option value="N"  >Night</option>
			    </select>
　班次：<input type="text" name="ShiftName" id="ShiftName" value="<%=ShiftName%>" size="2" readonly/>
				　<font color="red">JobNumber：<input type="text" name="JobNumber" id="JobNumber" value="<%=JobNumber%>" size="15" readonly/>
				</font></strong>
				</td>
			</tr>
			<tr>
				<td colspan="5" align="center">
					　Production Time：<input type="text" name="OpenTime" id="OpenTime" onChange="jsdt()" value="<%=OpenTime%>" size="5"/>
					　Down Time：<input type="text" name="StopTime" id="StopTime" value="<%=StopTime%>" size="5"/>
					　Blocked Products：<input type="text" name="PrBlocked" value="<%=PrBlocked%>" size="5"/>
				</td>
			</tr>
		
					  <tr>
						<th width="20%">Machine</th>
						<th width="20%">Good pieces</th>
						<th width="20%">Bad picesc</th>
						<th width="20%">Real FOR</th>
						<th width="20%">Target FOR</th>
					  </tr>
					  <tr>
						<td align="center">Transfer</td>
						<td align="center"><input type="text" name="T1G" id="T1G" onChange="jst1f()" value="<%=T1G%>" size="6" /></td>
						<td align="center"><input type="text" name="T1B" id="T1B" onChange="jst1f()" value="<%=T1B%>" size="6" /></td>
						<td align="center"><input type="text" name="T1F" id="T1F" value="<%=T1F%>" size="6" /></td>
						<td align="center"><input type="text" name="T1T" id="T1F" value="<%=T1T%>" size="6" style="background:#CCCCCC;" readonly /></td>
					  </tr>
						<tr>
						<td align="center">Cell 1</td>
						<td align="center"><input type="text" name="C1G" id="C1G" onChange="jscf(1)" value="<%=C1G%>" size="6" /></td>
						<td align="center"><input type="text" name="C1B" id="C1B" onChange="jscf(1)" value="<%=C1B%>" size="6" /></td>
						<td align="center"><input type="text" name="C1F" id="C1F" value="<%=C1F%>" size="6" /></td>
						<td align="center"><input type="text" name="C1T" id="C1T" value="<%=C1T%>" size="6" style="background:#CCCCCC;" readonly /></td>
					  </tr>
					
					  <tr>
						<td align="center">Cell 2</td>
						<td align="center"><input type="text" name="C2G" id="C2G" onChange="jscf(2)" value="<%=C2G%>" size="6" /></td>
						<td align="center"><input type="text" name="C2B" id="C2B" onChange="jscf(2)" value="<%=C2B%>" size="6" /></td>
						<td align="center"><input type="text" name="C2F" id="C2F" value="<%=C2F%>" size="6" /></td>
						<td align="center"><input type="text" name="C2T" id="C2T" value="<%=C2T%>" size="6" style="background:#CCCCCC;" readonly /></td>
					  </tr>
					
					  <tr>
						<td align="center">Cell 3</td>
						<td align="center"><input type="text" name="C3G" id="C3G" onChange="jscf(3)" value="<%=C3G%>" size="6" /></td>
						<td align="center"><input type="text" name="C3B" id="C3B" onChange="jscf(3)" value="<%=C3B%>" size="6" /></td>
						<td align="center"><input type="text" name="C3F" id="C3F" value="<%=C3F%>" size="6" /></td>
						<td align="center"><input type="text" name="C3T" id="C3T" value="<%=C3T%>" size="6" style="background:#CCCCCC;" readonly /></td>
					  </tr>
					<%if ProductName = "RA" or ProductName = "Slim" then%>
					  <tr>
						<td align="center">Cell 4</td>
						<td align="center"><input type="text" name="C4G" id="C4G" onChange="jscf(4)" value="<%=C4G%>" size="6" /></td>
						<td align="center"><input type="text" name="C4B" id="C4B" onChange="jscf(4)" value="<%=C4B%>" size="6" /></td>
						<td align="center"><input type="text" name="C4F" id="C4F" value="<%=C4F%>" size="6" /></td>
						<td align="center"><input type="text" name="C4T" id="C4T" value="<%=C4T%>" size="6" style="background:#CCCCCC;" readonly /></td>
					  </tr>
					
					  <tr>
						<td align="center">Cell 5</td>
						<td align="center"><input type="text" name="C5G" id="C5G" onChange="jscf(5)" value="<%=C5G%>" size="6" /></td>
						<td align="center"><input type="text" name="C5B" id="C5B" onChange="jscf(5)" value="<%=C5B%>" size="6" /></td>
						<td align="center"><input type="text" name="C5F" id="C5F" value="<%=C5F%>" size="6" /></td>
						<td align="center"><input type="text" name="C5T" id="C5T" value="<%=C5T%>" size="6" style="background:#CCCCCC;" readonly /></td>
					  </tr>
					  
					<%end if%>
		<tr><td colspan="5" align="center"><%=CellSubmit%></td></tr>
<%'END IF%>
</table>
</br>
<table width="1000" border="1" cellspacing="0" cellpadding="0">
	<tr>
		<td colspan="5" align="center">
			
		</td>
		</tr>
	</table>
</form>
</center>

</body> 
</html>