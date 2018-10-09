<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>HV Date Transfer</title>
<script type="text/javascript">
	window.opener=null;   
  	window.open("","_self");  
	window.close();
</script>
</head>
<body>
<%
set conn = Server.CreateObject("ADODB.Connection")
ConnString = "Provider=OraOLEDB.Oracle;PLSQLRSet=1;Password=welcome4;User ID=BAR_WEB_KEB;Data Source=KEBDB"
conn.Open ConnString

conn.execute("Delete from HVreport") '删除记录

NowDate = date()
OutDate = DateAdd("d",-1,date())
'response.write OutDate & "</br></br>"
set rs=server.createobject("adodb.recordset")
OraSql = "SELECT PRODUCTNAME,LINENAME,count(LineName) ShiftSum FROM report_hv_data RHD where ADATE = '"&OutDate&"' group by PRODUCTNAME,LINENAME"
'response.write OraSql & "</br></br>"
rs.open OraSql,conn,1,1
do while not rs.eof
	ProductName = rs("PRODUCTNAME")
	LINENAME = rs("LINENAME")
	ShiftSum = csng(rs("ShiftSum"))
	SET LoopRs = server.createobject("adodb.recordset")
	OraLoopSql = "SELECT PRODUCTNAME,LINENAME,SUM(NVL(T1G,0)) T1G,SUM(NVL(T1B,0)) T1B,SUM(NVL(C1G,0)) C1G,SUM(NVL(C1B,0)) C1B,SUM(NVL(C2G,0)) C2G,SUM(NVL(C2B,0)) C2B,SUM(NVL(C3G,0)) C3G,SUM(NVL(C3B,0)) C3B,SUM(NVL(C4G,0)) C4G,SUM(NVL(C4B,0)) C4B,SUM(NVL(C5G,0)) C5G,SUM(NVL(C5B,0)) C5B FROM report_hv_data RHD"
	OraLoopSql = OraLoopSql & " where ADATE = '"&OutDate&"' AND ProductName = '"&ProductName&"' AND LINENAME = '"&LineName&"' GROUP BY PRODUCTNAME,LINENAME"
	'response.write OraLoopSql & "</br></br>"
	LoopRs.open OraLoopSql,conn,1,1
	if not LoopRs.eof then
		set TarRs = server.createobject("adodb.recordset")
		OraTarSql = "select OutPCS,T1,C1,C2,C3,C4,C5 from report_hv_target"
		OraTarSql = OraTarSql & " WHERE ProductName = '"&ProductName&"' AND LineName = '"&LineName&"' and Tdate < = '"&OutDate&"' ORDER BY Tdate desc"
		'response.write PRODUCTNAME & " - " & LINENAME & "</br>"
		'response.write OraTarSql & "</br>"
		TarRs.open OraTarSql,conn,1,1
		OutPCS = 0
		T1T = 0
		C1T = 0
		C2T = 0
		C3T = 0
		C4T = 0
		C5T = 0
		if not TarRs.eof then
			OutPCS = csng(TarRs("OutPCS")) * ShiftSum
			'response.write TarRs("OutPCS") & " * "
			'response.write ShiftSum & " = "
			'response.write OutPCS & "</br>"
			
			T1T = TarRs("T1")
			C1T = TarRs("C1")
			C2T = TarRs("C2")
			C3T = TarRs("C3")
			C4T = TarRs("C4")
			C5T = TarRs("C5")
		end if
		TarRs.close
		Set TarRs = nothing

		'开始处理数据
		PRODUCT=trim(ProductName)
		IF PRODUCT = "RA" THEN PRODUCT = "RA 11x15 Speaker (Danubius)"
		LINE_NUM=trim(LineName)
		CREATE_TIME=trim(cstr(OutDate))
		LINE_TYPE=TODB("AUTO")
		TARGET=TODB(OutPCS)
		REAL_TO_WH=TODB("0")
		
		T1G = CSNG(LoopRs("T1G"))
		T1B = CSNG(LoopRs("T1B"))
		T1A = T1G + T1B
		IF T1A = 0 THEN
			T1F = 0
		ELSE
			T1F = T1G / T1A * 100
		END IF
		
		C1G = CSNG(LoopRs("C1G"))
		C1B = CSNG(LoopRs("C1B"))
		C1A = C1G + C1B
		IF C1A = 0 THEN
			C1F = 0
		ELSE
			C1F = C1G / C1A * 100
		END IF
		
		C2G = CSNG(LoopRs("C2G"))
		C2B = CSNG(LoopRs("C2B"))
		C2A = C2G + C2B
		IF C2A = 0 THEN
			C2F = 0
		ELSE
			C2F = C2G / C2A * 100
		END IF
		
		C3G = CSNG(LoopRs("C3G"))
		C3B = CSNG(LoopRs("C3B"))
		C3A = C3G + C3B
		IF C3A = 0 THEN
			C3F = 0
		ELSE
			C3F = C3G / C3A * 100
		END IF
		
		C4G = CSNG(LoopRs("C4G"))
		C4B = CSNG(LoopRs("C4B"))
		C4A = C4G + C4B
		IF C4A = 0 THEN
			C4F = 0
		ELSE
			C4F = C4G / C4A * 100
		END IF
		
		C5G = CSNG(LoopRs("C5G"))
		C5B = CSNG(LoopRs("C5B"))
		C5A = C5G + C5B
		IF C5A = 0 THEN
			C5F = 0
		ELSE
			C5F = C5G / C5A * 100
		END IF

		if PRODUCTNAME = "Petra" or PRODUCTNAME = "Franklin" THEN 'Petra or FrankLin
			ACTUAL=TODB(C3G)
			PQ_OUT=TODB(C2G)
			PQ_TARGET = (100 - T1T) * (100 - C1T) * (100 - C2T) / 10000
			PQ_YIELD = T1F * C1F * C2F / 10000
			AMS_GOOD = TODB(C3G)
			AMS_START = TODB(C3A)
			AMS_TARGET = TODB(C3T)
			AMS_YIELD = TODB(C3F)
			COSMETIC_GOOD = TODB(C3G)
			COSMETIC_START = TODB(C3G)
			TTY_OPEN = TODB(T1A)
			TTY_CLOSE = TODB(C3G)
			TTY_TARGET = (100 - C1T) * (100 - C2T) * (100 - C3T) / 10000
			TTY_YIELD = C1F * C2F * C3F / 10000
		else						'RA / Donau Slim
			ACTUAL=TODB(C5G)
			PQ_OUT=TODB(C4G)
			PQ_TARGET = (100 - T1T) * (100 - C1T) * (100 - C2T) * (100 - C3T)* (100 - C4T) / 100000000
			PQ_YIELD = T1F * C1F * C2F * C3F * C4F / 100000000
			AMS_GOOD = TODB(C5G)
			AMS_START = TODB(C5A)
			AMS_TARGET = TODB(C5T)
			AMS_YIELD = TODB(C5F)
			COSMETIC_GOOD=TODB(C5G)
			COSMETIC_START=TODB(C5G)
			TTY_OPEN=TODB(T1A)
			TTY_CLOSE=TODB(C5G)
			TTY_TARGET = (100 - C1T) * (100 - C2T) * (100 - C3T) * (100 - C4T) * (100 - C5T) / 100000000
			TTY_YIELD = C1F * C2F * C3F * C4F * C5F / 100000000
		end if
		
			PQ_START = TODB(T1A)
			AMS_TARGET = 100 - AMS_TARGET
			COSMETIC_TARGET = 100
			COSMETIC_YIELD = 100

			sql = "insert into HvReport (PRODUCT,LINE_NUM,CREATE_TIME,LINE_TYPE,TARGET,ACTUAL,REAL_TO_WH,PQ_OUT,PQ_START,PQ_TARGET,PQ_YIELD,AMS_GOOD,AMS_START,AMS_TARGET,AMS_YIELD,COSMETIC_GOOD,COSMETIC_START,COSMETIC_TARGET,COSMETIC_YIELD,TTY_OPEN,TTY_CLOSE,TTY_TARGET,TTY_YIELD)"
			sql = sql & " VALUES  ('"&PRODUCT&"','"&LINE_NUM&"','"&NowDate&"','"&LINE_TYPE&"','"&TARGET&"','"&ACTUAL&"','"&REAL_TO_WH&"','"&PQ_OUT&"','"&PQ_START&"','"&PQ_TARGET&"','"&PQ_YIELD&"','"&AMS_GOOD&"','"&AMS_START&"','"&AMS_TARGET&"','"&AMS_YIELD&"','"&COSMETIC_GOOD&"','"&COSMETIC_START&"','"&COSMETIC_TARGET&"','"&COSMETIC_YIELD&"','"&TTY_OPEN&"','"&TTY_CLOSE&"','"&TTY_TARGET&"','"&TTY_YIELD&"')"
			'response.write (sql) & "<br/>"
			conn.execute(sql)
			
			' response.write "PRODUCT: " & PRODUCT & "</br>"
			' response.write "LINE_NUM: " & LINE_NUM & "</br>"
			' response.write "T1G: " & T1G & "</br>"
			' response.write "T1B: " & T1B & "</br>"
			' response.write "T1A: " & T1A & "</br>"
			' response.write "T1F: " & T1F & "</br></br>"
			
			' response.write "C1G: " & C1G & "</br>"
			' response.write "C1B: " & C1B & "</br>"
			' response.write "C1A: " & C1A & "</br>"
			' response.write "C1F: " & C1F & "</br></br>"
			
			' response.write "C2G: " & C2G & "</br>"
			' response.write "C2B: " & C2B & "</br>"
			' response.write "C2A: " & C2A & "</br>"
			' response.write "C2F: " & C2F & "</br></br>"
			
			' response.write "C3G: " & C3G & "</br>"
			' response.write "C3B: " & C3B & "</br>"
			' response.write "C3A: " & C3A & "</br>"
			' response.write "C3F: " & C3F & "</br></br>"
			
			' response.write "C4G: " & C4G & "</br>"
			' response.write "C4B: " & C4B & "</br>"
			' response.write "C4A: " & C4A & "</br>"
			' response.write "C4F: " & C4F & "</br></br>"
			
			' response.write "C5G: " & C5G & "</br>"
			' response.write "C5B: " & C5B & "</br>"
			' response.write "C5A: " & C5A & "</br>"
			' response.write "C5F: " & C5F & "</br></br>"

			' response.write "PQ_TARGET: " & PQ_TARGET & "</br>"
			' response.write "PQ_YIELD: " & PQ_YIELD & "</br>"
			' response.write "AMS_GOOD: " & AMS_GOOD & "</br>"
			' response.write "AMS_START: " & AMS_START & "</br>"
			' response.write "TTY_TARGET: " & TTY_TARGET & "</br>"
			' response.write "TTY_YIELD: " & TTY_YIELD & "</br>"
			' response.write "</br>"
			
	end if
	LoopRs.close
	set LoopRs = nothing
rs.movenext
loop
rs.close
set rs = nothing

function TODB (TODBvalue)
	if isnull(TODBvalue) then
		TODB=""
	else
		TODB = TODBvalue
	end if
end function

%>
</body>
</html>