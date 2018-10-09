<script language="javascript">
var strSearch="Search Job|查询工单"
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|工单号" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|型号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchStatus="Status|状态" 
var arrSearchStatus=strSearchStatus.split("|")
var strSearchSheetNumber="Sheet Number|子工单编号" 
var arrSearchSheetNumber=strSearchSheetNumber.split("|")
var strSearchLineName="LineName|线别" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|工厂" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchElapsedTime="Elapsed Time|用时" 
var arrSearchElapsedTime=strSearchElapsedTime.split("|")
var strSearchM="m|分钟" 
var arrSearchM=strSearchM.split("|")
var strSearchJobStartTime="Job Start Time|工单开始时间" 
var arrSearchJobStartTime=strSearchJobStartTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="to|到" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Job list (Default in past 7 days)|浏览工单列表（默认7天以内）" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strStation="Station|站名" 
var arrStation=strStation.split("|")
var strOperator="Operator|操作工" 
var arrOperator=strOperator.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strPartType="Part Type|制程" 
var arrPartType=strPartType.split("|")
var strLine="Line|线别" 
var arrLine=strLine.split("|")
var strElapsedTime="Elapsed Time|用时" 
var arrElapsedTime=strElapsedTime.split("|")
var strStartTime="Start Time|开始时间" 
var arrStartTime=strStartTime.split("|")
var strCloseTime="Close Time|结束时间" 
var arrCloseTime=strCloseTime.split("|")
var strRecords="No Records|没有记录" 
var arrRecords=strRecords.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){} 
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchStatus.innerText=arrSearchStatus[<%=session("language")%>]}catch(e){}
try{inner_SearchSheetNumber.innerText=arrSearchSheetNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchLineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{inner_SearchFactory.innerText=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchElapsedTime.innerText=arrSearchElapsedTime[<%=session("language")%>]}catch(e){}
try{inner_SearchM.innerText=arrSearchM[<%=session("language")%>]}catch(e){}
try{inner_SearchCurrentStation.innerText=arrSearchCurrentStation[<%=session("language")%>]}catch(e){}
try{inner_SearchJobStartTime.innerText=arrSearchJobStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Operator.innerText=arrOperator[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_PartType.innerText=arrPartType[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_ElapsedTime.innerText=arrElapsedTime[<%=session("language")%>]}catch(e){}
try{inner_StartTime.innerText=arrStartTime[<%=session("language")%>]}catch(e){}
try{inner_CloseTime.innerText=arrCloseTime[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
}
language()
</script>