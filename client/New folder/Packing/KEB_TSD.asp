<!--#include virtual="/WOCF/TSDconn.asp" -->

<html><head><title></title>
<meta http-equiv="Content-Type" content="text/html; charsTSDet=gb2312">
<link href="../images/css.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="jquery-1.9.0.min.js"></script>
<script language="javascript" type="text/javascript">
function resetSession()
{
	//session.Abandon();
	//alert("TEST");
	location.replace("http://keb-dt056:8082/admin/KEB_TSD.asp?dd=qq");
	//location.href = "http://keb-dt056:8082/admin/KEB_PVS.asp?dd=qq";
	
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
dd=request("dd")
if dd="qq" then
 session("i")=0
end if 

if session("i")="" then
 session("i")=0
end if

response.Write("<script language='javascript' type='text/javascript'>")
			response.Write("$(document).ready(function(){")
			response.Write("$('#jsq').css({'font-size': '22pt', 'color':'#FF0000'});")
			response.Write(" $('#jsq').html('计数：');")
		    response.Write("$('#jsq').append(")
		    response.Write(session("i"))
		    response.Write(")")
			response.Write("});")

			response.Write("</script>")

%>

<table width="100%" border="0" align="center" cellpadding="3" cellspacing="1" bgcolor="#FFFFFF" class="tableBorder">
<tr>
<td height="26" align="center" background="../images/admin_bg_1.gif"><span class="STYLE6"><font color="#ffffff">信息查讯</font></span></td>
</tr>

</table>
<form name="form2" method="post" action="KEB_TSD.asp">
<table width="100%" align="center">
  <tr bgcolor="#E8F1FF"> 

<td width="74%" height="100"> 
  <table border="0" align="right" cellpadding="0" cellspacing="0">
<tr> 
  <td width="251" align="center"><span class="STYLE9">2D Code ：</span></td>
  <td width="549" height="126" align="center"><div align="left">
    <input name="Dcode" type="text" id="Dcode" onFocus="this.value=''" value="" size="50">
    
    <input type="submit" name="Submit2" value="检  查">
  </div></td>
</tr>
</table></td>
<td width="7%" align="center" > <input type="button" onClick="resetSession()" value="重新计数"/></td>

<td width="19%" id="jsq"></td>
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

if PNO="F7GW" then
response.Write("Jason semi")
elseif PNO="F44F" then
response.Write("GoldBerg HFL")
end if

response.Write("</DIV>")
response.Write("<BR>")


if  Dcode<>"" then	
					
					
					
					
'Dcode=trim(request("Dcode"))
DcodeL=len(Dcode)
BJD=left(Dcode,3)
'response.Write(DcodeL)
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
		

				

			
	


sql="select  * from package a ,serial_index b where a.serial_id=b.serial_id and serialnumber = '"&Dcode&"'" 
  		
		
			 rsTSD.open sql,connTSD,1,3
			
		   
				
  				if rsTSD.eof And rsTSD.bof then
       				Response.Write "<p align='center' class='contents' >"
					
					
					response.Write("<font size=20  color=#FF0000 >")
					 
						
					response.Write("2D Barcode PVS")
					
					
		 if  session("i")="" then
		 session("i")=1
		 else
		   session("i")=session("i")+1
		   end if
		  
		   response.Write(" 无数据")
		   
		   
		   response.Write("</font>")
		 
		   response.Write("</p>")
			
			response.Write("<script language='javascript' type='text/javascript'>")
			response.Write("$(document).ready(function(){")
			response.Write("$('#jsq').css({'font-size': '22pt', 'color':'#FF0000'});")
		    response.Write(" $('#jsq').html('计数：');")
		    response.Write("$('#jsq').append(")
		    response.Write(session("i"))
		    response.Write(")")
		
			response.Write("});")

			response.Write("</script>")

			
   				else
			
			response.Write("<script language='javascript' type='text/javascript'>")
			response.Write("$(document).ready(function(){")
			response.Write("$('#jsq').css({'font-size': '22pt', 'color':'#FF0000'});")
			response.Write(" $('#jsq').html('计数：');")
		    response.Write("$('#jsq').append(")
		    response.Write(session("i"))
		    response.Write(")")
			response.Write("});")

			response.Write("</script>")
				%>
				
<table class="tableBorder" width="100%" border="0" align="center" cellpadding="3" 

cellspacing="1" bgcolor="#FFFFFF">
<tr> 
<td colspan="9" align="center" background="../images/admin_bg_1.gif"><b><font 

color="#ffffff">查看与修改</font></b></td>
</tr>
<tr bgcolor="#E8F1FF"> 
  <td width="2%" height="61" align="center" class="STYLE5 STYLE8">&nbsp;</td>
  <td width="9%" align="center" class="STYLE6"><span class="STYLE9">PVS ID</span></td>
  <td width="7%" align="center" class="STYLE6"><span class="STYLE9">P/F</span></td>
  <td width="12%" align="center" class="STYLE6"><span class="STYLE9">ERROR</span></td>
  <td width="12%" align="center" class="STYLE6"><span class="STYLE9">CRITERION </span></td>
  <td width="13%" align="center" class="STYLE10">JOB Number </td>
  <td width="7%" align="center" class="STYLE10">TESTER</td>
  <td width="10%" align="center" class="STYLE10">PC  NAME</td>
  <td width="28%" align="center" class="STYLE6"><span class="STYLE9">PVS DATE TIME </span></td>
</tr>
<%do while not rsTSD.eof%>

<%
adfail=trim(rsTSD("testResult"))
cerrorname=rsTSD("failName")

if adfail="FAIL" then

  ' if  cerrorname="HOHD" or cerrorname="HOHD.2"  or cerrorname="HOHD.1"  or  cerrorname="HOHDmax1.4W"  or cerrorname="HOHDmax1.125W" or cerrorname="HOHD1" or cerrorname="HOHD2"  or cerrorname="HOHD 100-250Hz"  or cerrorname="HOHD 250-1kHz" or cerrorname="HOHDgap100-250Hz"  or cerrorname="HOHDgap250-1kHz"   or cerrorname="HOHDgap100-185Hz"or cerrorname="HOHDgap185-1kHz" then
   ' hohdA=rsTSD("hohd")
	%>
<tr bgcolor="FF0000"> 


<%else%>
<tr bgcolor="00CC00"> 
<%end if %>
<td height="72" align="center"><span class="STYLE8"></span></td>
<td ><span class="STYLE9"><%=rsTSD("serial_id")%></span></td>
<td align="center"><div align="center" class="STYLE9">
  <%
adfail=rsTSD("testresult")
response.Write(adfail)
'if adfail="Fail" then

'response.Write("Fail")

'else

'response.Write("PASS")

'end if

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
		
		set rsTSD=nothing
			
	
		
	
			
				end if	
				
	       end if
	
		dim endtime

endtime=timer()
	
		%>
	

</table>
<p>
页面执行时间：<font size="+1" color="#0000FF"><b><%=FormatNumber((endtime-startime),2,-1)%></b></font>秒</p>



</body>
</html>
