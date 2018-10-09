<!--#include file= "conn.asp"--> 
<!--#include file= "../include/function.asp"-->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<title>LPB</title>
</head>
<body>
<%
	'dim ProductName,LineName,LineNo,Hdate As String
	ProductName = request("ProductName")
	LineName = request("LineName")
	LineNo = ProductName&LineName
	Hdate = request("Hdate")
	HShift = request("HShift")
	DDN = request("DDN")
	HSSV = request("HSSV")
	HLE = request("HLE")
	Hespeed = request("Hespeed")
	Hahour = request("Hahour")
	Heqty = request("Heqty")
	Haqty = request("Haqty")
	'Hoee = request("Hoee")
	
	oee = Formatnumber((Haqty / (Hespeed * Hahour)) * 100,2)
	
	wima = request("wima")
	cell1 = request("cell1")
	cell2 = request("cell2")
	cell3 = request("cell3")
	cell4 = request("cell4")
	cell5 = request("cell5")
	
	sm = request("sm")
	oc = request("oc")
	ts = request("ts")
	mt = request("mt")
	
	Qablock = request("Qablock")
	Qarelease = request("Qarelease")
	
	Hsend = request("Hsend")
	Hback = request("Hback")

	'response.end
	'写入生产数量及废品率信息
	set rs = server.createobject("adodb.recordset") 
		sql = "select top 1 * FROM [Ministat].dbo.BJLPB2"
		rs.open sql,conn,1,3
			rs.addnew
			rs("HID") = "BU"
			rs("Hlineno") = LineNo
			rs("HDate") = Hdate
			rs("Udate") = now()
			rs("HD_N") = DDN
			rs("HShift") = HShift
			rs("HSSV") = HSSV
			rs("HLE") = HLE
			rs("Hespeed") = Hespeed
			rs("Hahour") = Hahour
			rs("Heqty") = Heqty
			rs("Haqty") = Haqty
			rs("HOEE") = oee
			rs("Wima") = wima
			rs("Cell1") = cell1
			rs("Cell2") = Cell2
			rs("Cell3") = Cell3
			rs("cell4") = cell4
			rs("Cell21") = Cell5
			rs.update
		rs.close
		set rs = nothing
		
		' Sqlstr1 = "INSERT INTO [Ministat].dbo.BJLPB2(HID,Hlineno,HDate,Udate,HD_N,HShift,HSSV,HLE,Hespeed,Hahour,Heqty,Haqty,HOEE,Wima,Cell1,Cell2,Cell3,cell4,Cell21)"
		' Sqlstr = Sqlstr1 & "VALUES('BU','"&LineNo&"','"+Hdate+"',GETDATE(),'"&DDN&"','"&HShift&"','"&HSSV&"','"&HLE&"','"+Hespeed+"','"+Hahour+"','"+Heqty+"','"+Haqty+"','"+oee+"','"&wima&"','"&cell1&"','"&Cell2&"','"&Cell3&"','"&cell4&"','"&Cell5&"')"
		' response.write Sqlstr
		' response.end
		
		
		'写入在线检测数据
		Sqlstr1 = "INSERT INTO [data-warehouse].dbo.OnlineData2(CID,CDate,Center,CLINE,Cshift,CShiftType,CType,CQTY)"
		Sqlstr = Sqlstr1 & " VALUES('CT','"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','MT','"&MT&"')"
		conn.execute(Sqlstr)
		Sqlstr = Sqlstr1 & " VALUES('CT','"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','OC','"&oc&"')"
		conn.execute(Sqlstr)
		Sqlstr = Sqlstr1 & " VALUES('CT','"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','SM','"&sm&"')"
		conn.execute(Sqlstr)
		Sqlstr = Sqlstr1 & " VALUES('CT','"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','TS','"&ts&"')"
		conn.execute(Sqlstr)
		'写入在线检测数据

		'写入封存数据
		Sqlstr1 = "INSERT INTO [Ministat].[dbo].[QABlock2](CDate,CCenter,CLINE,CShift,CShiftType,CType,CQTY,upddate)"
		Sqlstr = Sqlstr1 & " VALUES('"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Block','"&Qablock&"',GETDATE())"
		conn.execute(Sqlstr)
		Sqlstr = Sqlstr1 & " VALUES('"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Release','"&Qarelease&"',GETDATE())"
		conn.execute(Sqlstr)
		'写入封存数据

		Sqlstr1 = "INSERT INTO [Ministat].[dbo].[HGRecheck2](CDate,CCenter,CLINE,CShift,CShiftType,CType,CQTY,upddate)"
		Sqlstr = Sqlstr1 & " VALUES('"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Send','"&Hsend&"',GETDATE())"
		conn.execute(Sqlstr)
		Sqlstr = Sqlstr1 & " VALUES('"+Hdate+"','"&ProductName&"','"&LineNo&"','"&HShift&"','"&DDN&"','Back','"+Hback+"',GETDATE())"
		conn.execute(Sqlstr)

	call sussLoctionHref("添加成功","LpbOutput.asp")
%>
</body>
</html>
