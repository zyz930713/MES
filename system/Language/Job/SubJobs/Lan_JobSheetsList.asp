<script language="javascript">
var strSearch="Search Job|��ѯ����"
var arrSearch=strSearch.split("|")
var strSearchSheetNumber="Sheet Number|������" 
var arrSearchSheetNumber=strSearchSheetNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchJobType="Job Type|��������" 
var arrSearchJobType=strSearchJobType.split("|")
var strSearchLineName="LineName|�߱�" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|����" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchOptionFactory="Factory|����" 
var arrSearchOptionFactory=strSearchOptionFactory.split("|")
var strSearchStatus="Status|״̬" 
var arrSearchStatus=strSearchStatus.split("|")
var strSearchOptionStatus="All|����" 
var arrSearchOptionStatus=strSearchOptionStatus.split("|")
var strSearchCurrentStation="Current Station|��ǰվ" 
var arrSearchCurrentStation=strSearchCurrentStation.split("|")
var strSearchOptionCurrentStation="All|����" 
var arrSearchOptionCurrentStation=strSearchOptionCurrentStation.split("|")
var strSearchCycleTime="Cycle Time|��תʱ��" 
var arrSearchCycleTime=strSearchCycleTime.split("|")
var strBrowse="Browse Job Sheets list|��������б�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strStatus="Status|״̬" 
var arrStatus=strStatus.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strDelete="Delete|ɾ��" 
var arrDelete=strDelete.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strPartType="Part Type|�Ƴ�" 
var arrPartType=strPartType.split("|")
var strQuantity="Quantity|����" 
var arrQuantity=strQuantity.split("|")
var strLine="Line|�߱�" 
var arrLine=strLine.split("|")
var strStartTime="Start Time|��ʼʱ��" 
var arrStartTime=strStartTime.split("|")
var strCloseTime="Close Time|����ʱ��" 
var arrCloseTime=strCloseTime.split("|")
var strCycleTime="Cycle Time|��תʱ��" 
var arrCycleTime=strCycleTime.split("|")
var strStations="Included Stations|վ��" 
var arrStations=strStations.split("|")
var strOperators="Stations' Operators|������" 
var arrOperators=strOperators.split("|")
var strRecords="No Records|û�м�¼" 
var arrRecords=strRecords.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchSheetNumber.innerText=arrSearchSheetNumber[<%=session("language")%>]}catch(e){} 
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchJobType.innerText=arrSearchJobType[<%=session("language")%>]}catch(e){}
try{inner_SearchLineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{inner_SearchFactory.innerText=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrSearchOptionFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchStatus.innerText=arrSearchStatus[<%=session("language")%>]}catch(e){}
try{document.all.status.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{inner_SearchCurrentStation.innerText=arrSearchCurrentStation[<%=session("language")%>]}catch(e){}
try{document.all.currentstation.options[0].text=arrSearchOptionCurrentStation[<%=session("language")%>]}catch(e){}
try{inner_SearchJobStartTime.innerText=arrSearchJobStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchCycleTime.innerText=arrSearchCycleTime[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Delete.innerText=arrDelete[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_PartType.innerText=arrPartType[<%=session("language")%>]}catch(e){}
try{inner_Quantity.innerText=arrQuantity[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_StartTime.innerText=arrStartTime[<%=session("language")%>]}catch(e){}
try{inner_CloseTime.innerText=arrCloseTime[<%=session("language")%>]}catch(e){}
try{inner_CycleTime.innerText=arrCycleTime[<%=session("language")%>]}catch(e){}
try{inner_Stations.innerText=arrStations[<%=session("language")%>]}catch(e){}
try{inner_Operators.innerText=arrOperators[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
}
</script>