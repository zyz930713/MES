<script language="javascript">
var strBrowse="Browse a Job|�������" 
var arrBrowse=strBrowse.split("|")
var strSummary="Summary Info|ժҪ" 
var arrSummary=strSummary.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strPartType="Part Type|�Ƴ�" 
var arrPartType=strPartType.split("|")
var strStatus="Job Status|����״̬" 
var arrStatus=strStatus.split("|")
var strAction="Job Action|��������" 
var arrAction=strAction.split("|")
var strStartTime="Job Start Time|��ʼʱ��" 
var arrStartTime=strStartTime.split("|")
var strCloseTime="Job Close Time|����ʱ��" 
var arrCloseTime=strCloseTime.split("|")
var strStartQuantity="Start Quantity|��ʼ����" 
var arrStartQuantity=strStartQuantity.split("|")
var strGoodQuantity="Good Quantity|��Ʒ����" 
var arrGoodQuantity=strGoodQuantity.split("|")
var strDefectQuantity="Defect Quantity|ȱ������" 
var arrDefectQuantity=strDefectQuantity.split("|")
var strCycleTime="Cycle Time|��תʱ��" 
var arrCycleTime=strCycleTime.split("|")
var strDefectHistory="Defectcode Change History|ȱ�ݴ����޸���ʷ" 
var arrDefectHistory=strDefectHistory.split("|")
var strStationsInfo="Stations' Info|վ����Ϣ" 
var arrStationsInfo=strStationsInfo.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strUpdate="Update|����" 
var arrUpdate=strUpdate.split("|")
var strStation="Station Name|վ������" 
var arrStation=strStation.split("|")
var strOperator="Operator|������" 
var arrOperator=strOperator.split("|")
var strStationStartTime="Station Start Time|վ�㿪ʼʱ��" 
var arrStationStartTime=strStationStartTime.split("|")
var strStationCloseTime="Station Close Time|վ�����ʱ��" 
var arrStationCloseTime=strStationCloseTime.split("|")
var strElapsed="Elapsed|��ʱ" 
var arrElapsed=strElapsed.split("|")
var strMinus="Minus|��" 
var arrMinus=strMinus.split("|")
var strShiftOutTime="Shift Out Time|ͣ��ʱ��" 
var arrShiftOutTime=strShiftOutTime.split("|")
var strShiftInTime="Shift In Time|����ʱ��" 
var arrShiftInTime=strShiftInTime.split("|")
var strInterval="Interval|���" 
var arrInterval=strInterval.split("|")
var strEqual="Equal|����" 
var arrEqual=strEqual.split("|")
var strCycle="Cycle|��ת" 
var arrCycle=strCycle.split("|")
var strCheckAll="Check All|ȫ��ѡ��" 
var arrCheckAll=strCheckAll.split("|")
var strUncheckAll="Uncheck All|ȫ��ȡ��" 
var arrUncheckAll=strUncheckAll.split("|")
var strUpdate="Update|��  ��"
var arrUpdate=strUpdate.split("|")
var strReset="Reset|��  ��"
var arrReset=strReset.split("|")
function language()
{
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_Summary.innerText=arrSummary[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_PartType.innerText=arrPartType[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_StartTime.innerText=arrStartTime[<%=session("language")%>]}catch(e){}
try{inner_CloseTime.innerText=arrCloseTime[<%=session("language")%>]}catch(e){}
try{inner_StartQuantity.innerText=arrStartQuantity[<%=session("language")%>]}catch(e){}
try{inner_GoodQuantity.innerText=arrGoodQuantity[<%=session("language")%>]}catch(e){}
try{inner_DefectQuantity.innerText=arrDefectQuantity[<%=session("language")%>]}catch(e){}
try{inner_CycleTime.innerText=arrCycleTime[<%=session("language")%>]}catch(e){}
try{inner_DefectHistory.innerText=arrDefectHistory[<%=session("language")%>]}catch(e){}
try{inner_StationsInfo.innerText=arrStationsInfo[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Update.innerText=arrUpdate[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Operator.innerText=arrOperator[<%=session("language")%>]}catch(e){}
try{inner_StationStartTime.innerText=arrStationStartTime[<%=session("language")%>]}catch(e){}
try{inner_StationCloseTime.innerText=arrStationCloseTime[<%=session("language")%>]}catch(e){}
try{inner_Elapsed.innerText=arrElapsed[<%=session("language")%>]}catch(e){}
try{inner_Minus.innerText=arrMinus[<%=session("language")%>]}catch(e){}
try{inner_ShiftOutTime.innerText=arrShiftOutTime[<%=session("language")%>]}catch(e){}
try{inner_ShiftInTime.innerText=arrShiftInTime[<%=session("language")%>]}catch(e){}
try{inner_Interval.innerText=arrInterval[<%=session("language")%>]}catch(e){}
try{inner_Equal.innerText=arrEqual[<%=session("language")%>]}catch(e){}
try{inner_Cycle.innerText=arrCycle[<%=session("language")%>]}catch(e){}
try{document.all.CheckAll.value=arrCheckAll[<%=session("language")%>]}catch(e){}
try{document.all.UncheckAll.value=arrUncheckAll[<%=session("language")%>]}catch(e){}
try{document.all.Update.value=arrUpdate[<%=session("language")%>]}catch(e){}
try{document.all.Reset.value=arrReset[<%=session("language")%>]}catch(e){}
}
</script>