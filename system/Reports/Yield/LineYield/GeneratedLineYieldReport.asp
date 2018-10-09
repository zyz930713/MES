<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/Reports/Yield/YieldCheck.asp" -->
<!--#include virtual="/Functions/FormatFunctions.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/Functions/GetStatisticTestResult.asp" -->
<%
path=request.ServerVariables("PATH_INFO")
query=request.ServerVariables("QUERY_STRING")
query=replace(query,"&","*")
set rst=server.CreateObject("adodb.recordset")
jobnumber=trim(request("jobnumber"))
partnumber=trim(request("partnumber"))
machine_id=trim(request("machine_id"))
fromdate=request("fromdate")
todate=request("todate")
ordername=request("ordername")
ordertype=request("ordertype")
if ordername="" and ordertype="" then
order=" order by J.JOB_NUMBER desc"
else
order=" order by "&ordername&" "&ordertype
end if
where=""
if jobnumber<>"" then
where=where&" and J.JOB_NUMBER='"&jobnumber&"'"
end if
if partnumber<>"" then
where=where&" and P.PART_NUMBER='"&partnumber&"'"
end if
if machine_id<>"" then
where=where&" and j.MACHINE_ID='"&machine_id&"'"
end if
if fromdate<>"" then
where=where&" and J.TEST_TIME>=to_date('"&fromdate&"','yyyy-mm-dd')"
end if
if todate<>"" then
where=where&" and J.TEST_TIME<=todate('"&todate&"','yyyy-mm-dd')"
end if
pagepara="&jobnumber="&jobnumber&"&partnumber="&partnumber&"&jobstatus="&jobstatus&"&currentstation="&currentstation&"&fromdate="&fromdate&"&todate="&todate
pagename="/Reports/Yield/GeneratedStatisticTestYieldReport.asp"
SQL="select 1,JT.JOB_NUMBER from DISTINCT_JOB_TEST_RESULT JT inner join JOB J on JT.JOB_NUMBER=J.JOB_NUMBER where JT.JOB_NUMBER not like '%U%'"
	response.Write(SQL)
	response.End()
rs.open SQL,conn,1,3
if not rs.eof then
while not rs.eof
	G1=0
	G2=0
	G3=0
	D=0
	L=0
	H=0
	B=0
	F=0
	N=0
	R=0
	V=0
	I=0
	SUMALL=0	
	
	SQLt="select P.PART_PER_BOARD,P.PART_PER_FRAME from JOB J inner join PART P on J.PART_NUMBER_ID=P.NID where J.JOB_NUMBER='"&rs("JOB_NUMBER")&"'"
	rst.open SQLt,conn,1,3
	if not rst.eof then
	PPB=rst("PART_PER_BOARD")
	PPF=rst("PART_PER_FRAME")
	else
	end if
	rst.close
	
	SQLt="select * from JOB_TEST_RESULT where JOB_NUMBER='"&rs("JOB_NUMBER")&"' order by TEST_DEFECTCODE_ID"
	rst.open SQLt,conn,1,3
	while not rst.eof
	select case rst("TEST_DEFECTCODE_ID")
	case "TF00000001"
	G1=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000002"
	G2=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000003"
	G3=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000004"
	D=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000005"
	L=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000006"
	H=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000007"
	B=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000008"
	F=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000009"
	N=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000010"
	R=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000011"
	V=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000012"
	I=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000014" 'Frame
	FRAMES=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000015" 
	M_SEN=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000016" 
	M_NOISE=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000017"
	STD_N=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000018"
	STD_S=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000019"
	INK=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000020"
	ASS=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000021"
	TEST_START=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000022"
	TEST_GOOD=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000023"
	UNKNOWN=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000024"
	TEST_MODE=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000025"
	JOB_START=rst("TEST_DEFECTCODE_VALUE")
	case "TF00000028"
	TEST_YIELD=rst("TEST_DEFECTCODE_VALUE")
	end select
	rst.movenext
	wend
	rst.close

	SQLt="select TEST_DEFECTCODE_VALUE from JOB_TEST_RESULT where JOB_NUMBER='U"&rs("JOB_NUMBER")&"' and TEST_DEFECTCODE_ID='TF00000023'"'23 is "Unknown_Quantity"
	rst.open SQLt,conn,1,3
	U_UNKNOWN=rst("TEST_DEFECTCODE_VALUE")
	rst.close
	SQLt="select TEST_DEFECTCODE_VALUE from JOB_TEST_RESULT where JOB_NUMBER='U"&rs("JOB_NUMBER")&"' and TEST_DEFECTCODE_ID='TF00000025'"'25 is "Job Start Quantity"
	rst.open SQLt,conn,1,3
	U_JOB_START=rst("TEST_DEFECTCODE_VALUE")
	rst.close
	
	U_JOB=false
	SQLU="select * from JOB_TEST_RESULT where JOB_NUMBER='U"&rs("JOB_NUMBER")&"' order by TEST_TIME"
	rsU.open SQLU,conn,1,3
	if not rsU.eof then
	U_JOB=true
	While not rsU.eof
	
		select case rsU("TEST_DEFECTCODE_ID")
		case "TF00000001"
		G1=G1+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000002"
		G2=G2+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000003"
		G3=G3+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000004"
		D=D+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000005"
		L=L+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000006"
		H=H+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000007"
		B=B+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000008"
		F=F+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000009"
		N=N+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000010"
		R=R+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000011"
		V=V+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000012"
		I=I+rsU("TEST_DEFECTCODE_VALUE")
		case "TF00000014" 'Frame
		if U_UNKNOWN>UNKNOWN then
		FRAMES=rst("TEST_DEFECTCODE_VALUE")+(cint(U_UNKNOWN/PPB*4)+1)
		else
		FRAMES=0
		end if
		FRAMES=getmin(FRAMES,JOB_START/PPF)	
		case "TF00000021"'Test Start Quantity
		TEST_START=PPF*FRAMES
		case "TF00000023"
			if UNKNOWN>U_UNKNOWN then
				U=UNKNOWN-U_UNKNOWN
			else
				U=0
				gap=U_UNKNOWN-UNKNOWN
				if SUMALL>JOB_START then
					if G1-gap>0 then
					G1=G1-gap
					elseif G2-gap>0 then
					G2=G2-gap
					elseif G3-gap>0 then
					G3=G3-gap
					end if
				end if
			end if
		end select
	rsU.movenext
	wend
	end if
	rsU.close
	
	SUMALL=G1+G2+G3+D+L+H+B+F+N+R+V+I
	
	
		
	SQLTD = "select * from TEST_DEFECTCODE"
    rsTD.Open SQLTD, conn, 1, 3
    If Not rsTD.EOF Then
        SQLN="select * from JOB_TEST_STATISTIC_RESULT where JOB_NUMBER = '"&rs("JOB_NUMBER")&"'"
        rsN.Open SQLN, conn, 3, 3
        If rsN.EOF Then
        While Not rsTD.EOF
            Select Case rsTD("DEFECT_NAME")
            Case "Grade 1"
            defectcode_value = G1
            Case "Grade 2"
            defectcode_value = G2
            Case "Grade 3"
            defectcode_value = G3
            Case "Dead"
            defectcode_value = D
            Case "Low Sensitivity"
            defectcode_value = L
            Case "High Sensitivity"
            defectcode_value = H
            Case "Back Volume Leakage"
            defectcode_value = B
            Case "Front Volume Leakage"
            defectcode_value = F
            Case "Noise"
            defectcode_value = N
            Case "Raise"
            defectcode_value = R
            Case "Sensitivity Varriation"
            defectcode_value = V
            Case "Current"
            defectcode_value = I
            Case "Unkown"
            defectcode_value = U
            Case "Frame"
            defectcode_value = FRAMES
            Case "Test Yield"
            defectcode_value = TEST_YIELD
            Case "Mean Sensitivity"
            defectcode_value = M_SEN
            Case "Mean of Noise"
            defectcode_value = M_NOISE
            Case "STDDEV of Noise"
            defectcode_value = STD_N
            Case "STDDEV of Sensitivity"
            defectcode_value = STD_S
            Case "Ink Dot"
            defectcode_value = INK
            Case "Assembly Good Quantity"
            defectcode_value = ASS
            Case "Test Start Quantity"
            defectcode_value = TEST_START
            Case "Test Good Quantity"
            defectcode_value = TOTAL_GOOD
            Case "Test Total Quantity"
            defectcode_value = TOTAL_TEST
            Case "Test Mode"
            defectcode_value = TEST_MODE
            Case "Job Start Quantity"
            defectcode_value = JOB_START
            End Select
            
            rsN.AddNew
            rsN("JOB_NUMBER") = rs("JOB_NUMBER")
            rsN("MACHINE_ID") = rs("MACHINE_ID")
            rsN("TEST_DEFECTCODE_ID") = rsTD("NID")
            rsN("TEST_DEFECTCODE_VALUE") = defectcode_value
            rsN("TEST_TIME") = rs("TEST_TIME")
            rsN.Update
        rsTD.MoveNext
        Wend
        End If
        rsN.Close
        
    End If
    rsTD.Close

rs.movenext
wend
end if
rs.close

SQL="select 1,J.JOB_NUMBER,J.MACHINE_ID,J.TEST_TIME from JOB_TEST_STATISTIC_RESULT J where J.JOB_NUMBER is not null "&where&order
session("SQL")=SQL
rs.open SQL,conn,1,3
SQLT="select NID,DEFECT_NAME from TEST_DEFECTCODE order by NID"
rsT.open SQLT,conn,1,3
Tcount=rsT.recordcount+4
%>
<!--#include virtual="/Components/PageSelect.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>Barcode System</title>
<link href="/CSS/General.css" rel="stylesheet" type="text/css">
<link href="/CSS/dynCalendar.css" rel="stylesheet" type="text/css">
<script language=javascript src="/Components/sniffer.js" type=text/javascript></script>
<script language=javascript src="/Components/dynCalendar.js" type=text/javascript></script>
</head>

<body>
<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
  <tr>
    <td height="20" colspan="7" class="t-c-greenCopy"><span>Search Job </span></td>
  </tr>
  <tr>
    <td height="20"><span class="style1">Job Number</span> </td>
    <td height="20"><input name="jobnumber" type="text" id="jobnumber" value="<%=jobnumber%>"></td>
    <td>Part Number </td>
    <td><input name="partnumber" type="text" id="partnumber" value="<%=partnumber%>"></td>
    <td>Job Time</td>
    <td>From
      <input name="fromdate" type="text" id="fromdate2" value="<%=fromdate%>" size="10">
      <script language=JavaScript type=text/javascript>
	function calendar1Callback(date, month, year)
	{
	document.all.fromdate.value=year + '-' + month + '-' + date
	}
    calendar1 = new dynCalendar('calendar1', 'calendar1Callback');
                        </script>
&nbsp;to
<input name="todate" type="text" id="todate2" value="<%=todate%>" size="10">
<script language=JavaScript type=text/javascript>
	function calendar2Callback(date, month, year)
	{
	document.all.todate.value=year + '-' + month + '-' + date
	}
    calendar2 = new dynCalendar('calendar2', 'calendar2Callback');
                        </script>
&nbsp; </td>
    <td><img src="/Images/Find.gif" alt="Search" width="27" height="21" align="absmiddle" style="cursor:hand" onClick="javascript:document.form1.submit()"></td>
  </tr>
</table>

<table width="100%"  border="1" cellpadding="0" cellspacing="0" bordercolorlight="#000099" bordercolordark="#FFFFFF">
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy">Browse Test Yield </td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>" class="t-c-greenCopy"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="50%">User: <% =session("User") %></td>
        <td width="50%" valign="middle"><div align="right"><span style="cursor:hand" title="Export Data to EXCEL File" onClick="javascript:window.open('JobYield_Export.asp')">          <img src="/Images/EXCEL.gif" width="30" height="30"></span></div></td>
      </tr>
    </table></td>
</tr>
<tr>
  <td height="20" colspan="<%=Tcount%>"><!--#include virtual="/Components/PageSplit.asp" --></td>
  </tr>
<tr>
  <td height="20" class="t-t-Borrow"><div align="center">NO</div></td>
  <td class="t-t-Borrow"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Job Number<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.JOB_NUMBER&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></td>
  <td height="20" class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.MACHINE_ID&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Machiner<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.MACHINE_ID&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <td class="t-t-Borrow"><div align="center"><img src="/Images/Accend.gif" alt="Click to accend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.TEST_TIME&ordertype=asc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'">Test Time<img src="/Images/Decend.gif" alt="Click to decend records" width="10" height="10" align="absmiddle" style="cursor:hand" onClick="javascript:location.href='<%=pagename%>?ordername=J.TEST_TIME&ordertype=desc&pagenum=<%=firstpage%>&pagesize_a=<%=pagesize_s%><%=pagepara%>'"></div></td>
  <%
  if not rsT.eof then
  While not rsT.eof%>
  <td class="t-t-Borrow"><div align="center"><%=rsT("DEFECT_NAME")%></div></td>
  <%
  rsT.movenext
  wend
  rsT.movefirst
  end if%>
  </tr>
<%
i=1
if not rs.eof then
rs.absolutepage=session("strpagenum")
while not rs.eof and i<=rs.pagesize%>
<tr>
  <td height="20"><div align="center"><% =(session("strpagenum")-1)*pagesize_s+i%></div></td>
    <td><div align="center"><%= rs("JOB_NUMBER") %></div></td>
    <td height="20"><div align="center"><%= rs("MACHINE_ID") %></div></td>
	 <td><div align="center"><% =formatdate(rs("TEST_TIME"),application("shortdateformat")) %></div></td>
	 <%
	  if not rsT.eof then
	  While not rsT.eof%>
	  <div align="center"><td><div align="center"><%=getStatisticTestResult(rs("JOB_NUMBER"),rsT("NID"))%>&nbsp;</div></td></div>
	  <%
	  rsT.movenext
	  wend
	  rsT.movefirst
	  end if
	  %>
  </tr>
<%
i=i+1
rs.movenext
wend
else
%>
  <tr>
    <td height="20" colspan="5"><div align="center">No Records </div></td>
  </tr>
<%end if
rsT.close
rs.close
set rsT=nothing%>
</table>
<!--#include virtual="/Components/CopyRight.asp" -->
</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->