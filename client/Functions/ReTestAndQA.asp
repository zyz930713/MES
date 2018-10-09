<%
	'Get Current AQL
	function GetCurrentAQL(ModelName)
		'判断是否是新的型号,如果是(3个月以内的),AQL=0.15
		SQL="SELECT MIN(START_TIME) FROM JOB WHERE PART_NUMBER_TAG='"&ModelName&"'"
		set rsIsNewModel=server.createobject("adodb.recordset")
		rsIsNewModel.open SQL,conn,1,3
		if rsIsNewModel.recordcount>0 then
			DiffTime=datediff("d",rsIsNewModel(0),now())
			if (DiffTime<90) then
				response.write "<script>window.alert('"+ModelName+"是一个新的型号,AQL为0.15')</script>"
				GetCurrentAQL="0.15"
				exit function
			end if
		end if
		'如果不是新型号,AQL参照设定值
		SQL="SELECT * FROM GENERAL_SETTING WHERE MODELNAME='"&ModelName&"'"
		set rsGeneralSetting=server.createobject("adodb.recordset")
		rsGeneralSetting.open SQL,conn,1,3
		if rsGeneralSetting.recordcount>0 then
			GetCurrentAQL=rsGeneralSetting("CURRENTAQL")
		else
			GetCurrentAQL=""
		end if
	end function
	
	'Set Current AQL
	function SetCurrentAQL(ModelName,LastTestResult,LastTestAQL,TestBatchNoList,LastFailUnit)
		'Add by jack zhang 2011-2-25
		SQL="INSERT INTO QA_TESTHISTORY_BAK(ModelName,TestDateTime,TestAQL,TestResult,TestBatchNoList,RESULT)"
		SQL=SQL+"VALUES('"&ModelName&"',sysdate,'"&LastTestAQL&"','"&LastTestResult&"','"&TestBatchNoList&"','"+LastTestResult+"')"
		set rsBAK=server.createobject("adodb.recordset")
		rsBAK.open SQL,conn,1,3
		'end add
		
	if LastTestResult="PASS" then
			SQL="INSERT INTO QA_TESTHISTORY(ModelName,TestDateTime,TestAQL,TestResult,TestBatchNoList)"
			SQL=SQL+"VALUES('"&ModelName&"',sysdate,'"&LastTestAQL&"','"&LastTestResult&"','"&TestBatchNoList&"')"			
			set rs0=server.createobject("adodb.recordset")
			rs0.open SQL,conn,1,3
			
			SQL="SELECT * FROM QA_TESTHISTORY WHERE MODELNAME='"&ModelName&"'"
			set rsPASS=server.createobject("adodb.recordset")
			rsPASS.open SQL,conn,1,3
			if rsPASS.recordcount>=5 AND cdbl(LastTestAQL)=0.4 then
				SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.65 WHERE MODELNAME='"&ModelName&"'"
				set rsUpdate1=server.createobject("adodb.recordset")
				rsUpdate1.open SQL,conn,1,3
				
				
				SQL="DELETE QA_TESTHISTORY WHERE MODELNAME='"&ModelName&"'"
				set rsUpdate0=server.createobject("adodb.recordset")
				rsUpdate0.open SQL,conn,1,3
			
			
				SetCurrentAQL=0.65
			end if
				
			if rsPASS.recordcount>=10 AND cdbl(LastTestAQL)=0.65 then
				SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=1 WHERE MODELNAME='"&ModelName&"'"
				set rsUpdate2=server.createobject("adodb.recordset")
				rsUpdate2.open SQL,conn,1,3
				
				SQL="DELETE QA_TESTHISTORY WHERE MODELNAME='"&ModelName&"'"
				set rsUpdate3=server.createobject("adodb.recordset")
				rsUpdate3.open SQL,conn,1,3
				
				
				SetCurrentAQL=1
			end if
				
			if cdbl(LastTestAQL)=1 then
				SetCurrentAQL=1
			end if
			
		end if 
		
		if LastTestResult="FAIL" then
				SQL="DELETE QA_TESTHISTORY WHERE MODELNAME='"&ModelName&"'"
				set rsUpdate5=server.createobject("adodb.recordset")
				rsUpdate5.open SQL,conn,1,3
	
				if cdbl(LastTestAQL)=0.4 then
					SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.4 WHERE MODELNAME='"&ModelName&"'"
					set rsUpdate6=server.createobject("adodb.recordset")
					rsUpdate6.open SQL,conn,1,3
					SetCurrentAQL=0.4
				end if
				
				if cdbl(LastTestAQL)=0.65 then
					SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.4 WHERE MODELNAME='"&ModelName&"'"
					set rsUpdate7=server.createobject("adodb.recordset")
					rsUpdate7.open SQL,conn,1,3
					SetCurrentAQL=0.4
				end if
				
				if cdbl(LastTestAQL)=1 then
					SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.65 WHERE MODELNAME='"&ModelName&"'"
					set rsUpdate9=server.createobject("adodb.recordset")
					rsUpdate9.open SQL,conn,1,3
					SetCurrentAQL=0.65
				end if
		end if 
	end function	

	'Set Current AQL
	function SetCurrentAQL_BAK(ModelName,LastTestResult,LastTestAQL,TestBatchNoList,LastFailUnit)
		'response.write LastTestResult &"<br>"
		'response.write  LastTestAQL &"<br>"
		if LastTestResult="PASS" then
			SQL="INSERT INTO QA_TESTHISTORY(ModelName,TestDateTime,TestAQL,TestResult,TestBatchNoList)"
			SQL=SQL+"VALUES('"&ModelName&"',sysdate,'"&LastTestAQL&"','"&LastTestResult&"','"&TestBatchNoList&"')"
			set rs0=server.createobject("adodb.recordset")
			rs0.open SQL,conn,1,3
			'response.write SQL
			SQL="SELECT COUNT(*) FROM QA_TESTHISTORY WHERE MODELNAME='"&ModelName&"'"
			set rsPASS=server.createobject("adodb.recordset")
			rsPASS.open SQL,conn,1,3
			if cint(rsPASS(0))>=10 then
				if cdbl(LastTestAQL)=0.4 then
					SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.65 WHERE MODELNAME='"&ModelName&"'"
					set rsUpdate1=server.createobject("adodb.recordset")
					rsUpdate1.open SQL,conn,1,3
					SetCurrentAQL=0.65
				end if
				
				if cdbl(LastTestAQL)=0.65 then
					SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=1 WHERE MODELNAME='"&ModelName&"'"
					set rsUpdate2=server.createobject("adodb.recordset")
					rsUpdate2.open SQL,conn,1,3
					SetCurrentAQL=1
				end if
				
				if cdbl(LastTestAQL)=1 then
				'	SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=1 WHERE MODELNAME='"&ModelName&"'"
'					set rsUpdate2=server.createobject("adodb.recordset")
'					rsUpdate2.open SQL,conn,1,3
					SetCurrentAQL=1
				end if
				
				SQL="DELETE QA_TESTHISTORY WHERE MODELNAME='"&ModelName&"'"
				set rsUpdate0=server.createobject("adodb.recordset")
				rsUpdate0.open SQL,conn,1,3
			else
				SetCurrentAQL=LastTestAQL
			end if
		end if 
		
		if LastTestResult="FAIL" then
		
				SQL="DELETE QA_TESTHISTORY WHERE MODELNAME='"&ModelName&"'"
				set rsUpdate5=server.createobject("adodb.recordset")
				rsUpdate5.open SQL,conn,1,3
				
				if (LastFailUnit>=2) then
					SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.4 WHERE MODELNAME='"&ModelName&"'"
					set rsUpdate8=server.createobject("adodb.recordset")
					rsUpdate8.open SQL,conn,1,3
					SetCurrentAQL=0.4
						
				else				
					if cdbl(LastTestAQL)=0.4 then
						SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.4 WHERE MODELNAME='"&ModelName&"'"
						set rsUpdate6=server.createobject("adodb.recordset")
						rsUpdate6.open SQL,conn,1,3
						SetCurrentAQL=0.4
					end if
					
					if cdbl(LastTestAQL)=0.65 then
						SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.4 WHERE MODELNAME='"&ModelName&"'"
						set rsUpdate7=server.createobject("adodb.recordset")
						rsUpdate7.open SQL,conn,1,3
						SetCurrentAQL=0.4
					end if
					
					if cdbl(LastTestAQL)=1 then
						SQL="UPDATE GENERAL_SETTING SET CURRENTAQL=0.65 WHERE MODELNAME='"&ModelName&"'"
						set rsUpdate9=server.createobject("adodb.recordset")
						rsUpdate9.open SQL,conn,1,3
						SetCurrentAQL=0.65
					end if
				end if 
		end if 
	end function	
	
	
	'Get QA_AQL
	function QA_SampleQty(TotalQty,AQL)
		SQL="SELECT SAMPLEQTY FROM QA_AQL WHERE QTYFROM<="&TotalQty&" AND QTYTO>="&TotalQty&" AND AQL="&AQL
		set rsSampleQty=server.createobject("adodb.recordset")
	 
		rsSampleQty.open SQL,conn,1,3
		if rsSampleQty.recordcount>0 then
			SampleQty=rsSampleQty("SAMPLEQTY").value
			if(SampleQty="100%")then
				QA_SampleQty=TotalQty
			else
				QA_SampleQty=SampleQty
			end if
		else
			QA_SampleQty=""
		end if
		rsSampleQty.close
	end function
	
	'Get QA_AQL
	function GetSelectAQL(TotalQty,selAQL)
		strOut=""
		if TotalQty <> "" then
			temQty=0
			SQL="SELECT AQL,SAMPLEQTY FROM QA_AQL WHERE QTYFROM<="&TotalQty&" AND QTYTO>="&TotalQty&" ORDER BY AQL"
			set rsAQL=server.createobject("adodb.recordset")	 
			rsAQL.open SQL,conn,1,3
			while not rsAQL.eof
				if rsAQL("SAMPLEQTY")="100%" then
					temQty=TotalQty
				else
					temQty=rsAQL("SAMPLEQTY")	
				end if
				strOut=strOut&"<option value='"&rsAQL("AQL")&"-"&temQty&"'"
				if selAQL=rsAQL("AQL") then
					strOut=strOut&" selected "
				end if
				strOut=strOut&">"&rsAQL("AQL")&"</option>"
				rsAQL.movenext
			wend
			rsAQL.close
		end if
		GetSelectAQL=strOut
	end function	
%>