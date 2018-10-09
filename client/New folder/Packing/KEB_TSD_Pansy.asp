<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<!--#include virtual="/WOCF/TSDconn.asp" -->

<html><head><title></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">

<script type="text/javascript" src="../Scripts/jquery-1.8.3.js"></script>
<script type="text/javascript" src="../Scripts/jquery-ui-1.9.2/js/jquery-ui-1.9.2.custom.min.js"></script>
<script language="javascript" type="text/javascript">
function resetSession()
{
	
	location.replace("http://10.6.100.57:9810/packing/KEB_TSD_Pansy.asp?JSQ=zero");
	
	
}
</script>

<style type="text/css">
<!--
.STYLE5 {font-size: 18pt}
.STYLE6 {
	font-size: 14pt;
	font-weight: bold;
}
.STYLE7 {
	font-size: larger;
	font-family: Arial, Helvetica, sans-serif;
}
.STYLE8 {font-family: Arial, Helvetica, sans-serif}
.STYLE9 {font-size: 18pt; font-family: Arial, Helvetica, sans-serif; }
.STYLE10 {font-size: 14pt; font-weight: bold; font-family: Arial, Helvetica, sans-serif; }
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
}
.RED{font-size: 22pt; font-family: Arial, Helvetica, sans-serif; }
.STYLE11 {font-size: x-large}

-->
</style>
</head>
<body onload= "javascript:document.all.Dcode.focus(); ">


<%

JSQ=request("JSQ")
if JSQ="zero" then
 session(("CheckA"))=0
 session(("CheckB"))=0
 response.Redirect("http://10.6.100.57:9810/packing/KEB_TSD_Pansy.asp")
end if 

if session(("CheckA"))="" then
 session(("CheckA"))=0
end if

if session(("CheckB"))="" then
 session(("CheckB"))=0
end if



%>
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF" class="tableBorder">
<tr>
<td height="26" align="center" background="../images/admin_bg_1.gif"><span class="STYLE6"><font color="#ffffff">信息查讯</font></span></td>
</tr>

</table>
<form name="form2" method="post" action="">
<table width="100%" align="center">
  <tr bgcolor="#E8F1FF"> 

<td width="74%" height="100"> 
  <table border="0" align="right" cellpadding="0" cellspacing="0">
<tr> 
  <td width="251" align="center"><span class="RED">2D Code ：</span></td>
  <td width="549" height="126" align="center"><div align="left">
    <input name="Dcode" type="text" id="Dcode" onFocus="this.value=''" value="" size="50">
    
    <input type="submit" name="Submit2" value="检  查">
  </div></td>
</tr>
</table></td>
<td width="7%" align="center" ><input type="button" onClick="resetSession()" value="重新计数"/> </td>

<td width="19%"><table width="200" border="1">
  <tr>
    <td id="CheckA">CheckA</td>
  </tr>
  <tr>
    <td id="CheckB">CheckB</td>
  </tr>
  
</table></td>
  </tr></table>
</form>





				
				
<%

dim startime
startime=timer()
Dcode=trim(request("Dcode"))
PNO=mid(Dcode,12,4) 
response.Write("<BR>")
response.Write("<DIV align=center class=STYLE5>")
DcodeA=left(Dcode,7)
DcodeB=mid(Dcode,8,1)
DcodeC=mid(Dcode,9,9)
response.Write(DcodeA)
response.Write("<font size=18>")
response.Write(DcodeB)
response.Write("</font>")
response.Write(DcodeC)
response.Write("&nbsp;&nbsp;&nbsp;")
response.Write("</DIV>")
response.Write("<BR>")
if  Dcode<>"" then	
DcodeL=len(Dcode)
BJD=left(Dcode,3)
PNO=mid(Dcode,12,4) 
Ver=mid(Dcode,16,1)
JSSYearA=mid(Dcode,4,1)
JSSWeekA=mid(Dcode,5,2)
JSSDayA=mid(Dcode,7,1)
line=mid(Dcode,8,1)
if DcodeL<>"17" then
response.Write("2D Barcode 位数不对！")
response.End()
end if
		

				

			

			
'sql="SELECT ad_id, linename, measuredatetime,AMS_line_name,AMSdatetime, adfail, error_name, cerror_name, pvs.func_GetHOHD(cerror_name) AS hohd, measurementpcname, preassemblycode FROM dbo.vw_adid_by_sn_new WHERE  ad_id IN (select ad_id from pvs.ad_serial where serialnumber = '"&Dcode&"') ORDER BY measuredatetime DESC"

sql="select  * from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&Dcode&"'" 

		
rsTSD.open sql,connTSD,1,1

if rsTSD.eof And rsTSD.bof then
			
else
			



sqlV="select SERIAL_NUMBER,TEST_TIME,TEST_NAME,TEST_MARGIN from  [TSD_Package].[dbo].[TEST_RESULT_DETAIL_S]   where  TEST_NAME   in ( 'FRES')  and  SERIAL_NUMBER='"&Dcode&"'"
	
	rsA.open sqlV,connTSD,1,3
	do while not rsA.eof
		TEST_NAME=rsA("TEST_NAME")		
		response.Write("<div align='center' ><font color='red'>"&TEST_NAME&"</font>")
		response.Write("&nbsp;&nbsp;")
	    response.Write(rsA("TEST_MARGIN"))
	    response.Write("<BR><BR></div>")
		

		sqlC="select * from CHECKABCSOL where paramSF='"&TEST_NAME&"' and Product_Name='Pansy'"
	
	    rs.open sqlC,conn,1,3
		  
		if not(rs.eof And rs.bof) then	
		
			YY=abc(CSng (rsA("TEST_MARGIN")),CSng (rs("ULA")),CSng (rs("LLA")),CDbl(rs("ULB")),CSng (rs("LLB")))
			if YY ="A" then
			AA=cint(AA)+1
			elseif YY="B" then
			BB=cint(BB)+1
			else
			CC=cint(CC)+1
			end if
		end if
		rs.close

	rsA.movenext
	loop
    rsA.close


function abc(tvalue,UA,LA,UB,LB)
	if tvalue < UA and tvalue > LA then
	      abc="A"
	elseif tvalue < UB  and tvalue >= LB then
		  abc= "B"
	else
		  abc="C" 
	end if
end function

%>
<div align="center" ><font size="19">
<%


	if CC<>0  then
		codetype="C"
		response.Write("<font color='red'>"+"C"+"类"+"</font>")
	elseif  BB<>0  then
		codetype="B"
		response.Write("<font color='Yellow'>"+"B"+"类"+"</font>") 
		
		
		if  session(("CheckB"))="" then
		 session(("CheckB"))=1
		 else
		 session(("CheckB"))=session(("CheckB"))+1
		 end if
		
		    response.Write("<script language='javascript' type='text/javascript'>")
			response.Write("$(document).ready(function(){")
			response.Write("$('#CheckA').css({'font-size': '22pt', 'color':'#00FF00'});")
		    response.Write(" $('#CheckA').html('CheckA：');")
		    response.Write("$('#CheckA').append(")
		    response.Write(session(("CheckA")))
		    response.Write(")")
		    response.Write("});")
            response.Write("</script>")
		
			response.Write("<script language='javascript' type='text/javascript'>")
			response.Write("$(document).ready(function(){")
			response.Write("$('#CheckB').css({'font-size': '22pt', 'color':'#FFFF00'});")
		    response.Write(" $('#CheckB').html('CheckB：');")
		    response.Write("$('#CheckB').append(")
		    response.Write(session(("CheckB")))
		    response.Write(")")
			response.Write("});")
			response.Write("</script>")			
		
		
	elseif  AA<>0  then
		codetype="A"
		response.Write("<font color='Greed'>"+"A"+"类"+"</font>")
		
		if  session(("CheckA"))="" then
		 session(("CheckA"))=1
		 else
		 session(("CheckA"))=session(("CheckA"))+1
		 end if
			
		    response.Write("<script language='javascript' type='text/javascript'>")
			response.Write("$(document).ready(function(){")
			response.Write("$('#CheckA').css({'font-size': '22pt', 'color':'#00FF00'});")
		    response.Write(" $('#CheckA').html('CheckA：');")
		    response.Write("$('#CheckA').append(")
		    response.Write(session(("CheckA")))
		    response.Write(")")
		    response.Write("});")
            response.Write("</script>")
		
			response.Write("<script language='javascript' type='text/javascript'>")
			response.Write("$(document).ready(function(){")
			response.Write("$('#CheckB').css({'font-size': '22pt', 'color':'#FFFF00'});")
		    response.Write(" $('#CheckB').html('CheckB：');")
		    response.Write("$('#CheckB').append(")
		    response.Write(session(("CheckB")))
		    response.Write(")")
			response.Write("});")
			response.Write("</script>")			
		
		
	end if
	
	
	 set rsJ=server.CreateObject("adodb.recordset")
     sqlJ= "select code from CHECKABCTYPEJSQ where code ='"&Dcode&"'" 
	 rsJ.Open sqlJ,conn,1,3
     if rsJ.bof and rsJ.eof then
	  CodeDate=now()
	  PRODUCT_NAME="Pansy"
	   conn.Execute("INSERT INTO CHECKABCTYPEJSQ (code,codetype,PRODUCT_NAME,CHECKABCDATE) values ('" &Dcode&"','" &codetype&"','" &PRODUCT_NAME&"','"&CodeDate&"')")	
		
	end if
	rsj.close
	
	
	%>
</font></div>
	<BR>
	
	
	
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" 

cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="11" align="center" background="../images/admin_bg_1.gif"><b><font 

color="#ffffff">查看与修改</font></b></td>
</tr>
<tr bgcolor="#E8F1FF"> 
  <td width="2%" height="61" align="center" class="STYLE5 STYLE8">&nbsp;</td>
  <td width="9%" align="center" class="STYLE6"><span class="STYLE9">TSD ID</span></td>
  <td width="7%" align="center" class="STYLE6"><span class="STYLE9">P/F</span></td>
  <td width="12%" align="center" class="STYLE6"><span class="STYLE9">ERROR</span></td>
  <td width="12%" align="center" class="STYLE6"><span class="STYLE9">CRITERION </span></td>
  <td width="13%" align="center" class="STYLE10">JOB Number </td>
  <td width="7%" align="center" class="STYLE10">TESTER</td>
  <td width="10%" align="center" class="STYLE10">PC  NAME</td>
  <td width="28%" align="center" class="STYLE6"><span class="STYLE9">TSD DATE TIME </span></td>
</tr>
<%do while not rsTSD.eof%>

<%
adfail=trim(rsTSD("testResult"))
cerrorname=rsTSD("failName")

if adfail="FAIL" then

  ' if  cerrorname="HOHD" or cerrorname="HOHD.2"  or cerrorname="HOHD.1"  or  cerrorname="HOHDmax1.4W"  or cerrorname="HOHDmax1.125W" or cerrorname="HOHD1" or cerrorname="HOHD2"  or cerrorname="HOHD 100-250Hz"  or cerrorname="HOHD 250-1kHz" or cerrorname="HOHDgap100-250Hz"  or cerrorname="HOHDgap250-1kHz"   or cerrorname="HOHDgap100-185Hz"or cerrorname="HOHDgap185-1kHz" then
   ' hohdA=rs("hohd")
	%>
<tr bgcolor="FF0000"> 


<%else%>
<tr bgcolor="00CC00"> 
<%end if %>
<td height="72" align="center"><span class="STYLE8"></span></td>
<td ><span class="STYLE9"><%=rsTSD("serial_id")%></span></td>
<td align="center"><div align="center" class="STYLE9">
  <%
adfail=rsTSD("testResult")
response.Write(adfail)


%>
</div></td>
<td align="center"><div align="center" class="STYLE9">
<%=rsTSD("failName")%>

</div>  </td>




 
<td align="center">&nbsp;</td>
<td align="center" class="STYLE9"><%=rsTSD("jobNumber")%></td>
<td align="center" class="STYLE9"><%=rsTSD("testitem")%></td>
<td align="center" class="STYLE9"><%=rsTSD("pcName")%></td>
<td align="center" class="STYLE5"><%=rsTSD("testTime")%></td>
</tr>
          <%
		
		
	
		rsTSD.movenext
		loop
		 a=a+1
		' response.Write(a)
		rsTSD.close
		set rsA=nothing
		
		set rs=nothing
			
	
		
	
			
				end if	
				
	       end if
	
		dim endtime

endtime=timer()
	
		%>
</table>
 



</body>
</html>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<!--#include virtual="/WOCF/TSD_Close.asp" -->