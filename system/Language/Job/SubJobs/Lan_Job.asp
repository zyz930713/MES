<script language="javascript">
var strSearch="Search Job|��ѯ����"
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchJobType="Job Type|��������" 
var arrSearchJobType=strSearchJobType.split("|")
var strSearchLineName="Line Name|�߱�" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|����" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchFamily="Family|����" 
var arrSearchFamily=strSearchFamily.split("|")
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
var strSearchJobStartTime="Start Time|��ʼʱ��" 
var arrSearchJobStartTime=strSearchJobStartTime.split("|")
var strSearchJobCloseTime="Close Time|�ر�ʱ��" 
var arrSearchJobCloseTime=strSearchJobCloseTime.split("|")
var strSearchCycleTime="Cycle Time|��תʱ��" 
var arrSearchCycleTime=strSearchCycleTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Sub Job list (Default in past 7 days)|����ӹ����б�Ĭ��7�����ڣ�" 
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

var strCurrentStations="Current Stations|��ǰվ��" 
var arrCurrentStations=strCurrentStations.split("|")

var strOperators="Stations' Operators|������" 
var arrOperators=strOperators.split("|")
var strLineLost="Line Lost|������ʧ" 
var arrLineLost=strLineLost.split("|")
var strRecords="No Records|û�м�¼" 
var arrRecords=strRecords.split("|")
var arrHoldReport=["Hold Report","����Hold����"]
var arrHoldTime=["Hold Time","��ͣʱ��"]
var arrHoldType=["Hold Type","��ͣ���"]
var arrHoldReason=["Hold Reason","��ͣԭ��"]
var arr2DCodeReport=["2D Code Report","��ά�뱨��"]
var arr2DCode=["2D Code","��ά��"]
var arrLinkTime=["Link Time","��ʱ��"]
var arrLinkUser=["Link User","����Ա"]
var arrSheetNumber=["Sheet Number","������"]
var arrReleaseTIme=["Release Time","�ͷ�ʱ��"]
var arrPSSTIme=["PSS Time","PSSʱ��"]
var arrHoldPerson=["Hold Person","��ͣ��Ա"]
var arrReleasePerson=["Release Person","�ͷ���Ա"]
var arrOpCode=["Op Code","������"]
var arrHoldStation=["Hold Station","��ͣվ��"]
var arrHoldJob=["Hold Job","��ͣ����"]
var arrReleaseJob=["Release Job","�ͷŹ���"]
var arrReleaseReason=["Release Reason","�ͷ�ԭ��"]
var arrJobProductRecord=["Job Production Records","����������¼"]
function language()
{
try{inner_JobProductRecord.innerText=arrJobProductRecord[<%=session("language")%>]}catch(e){} 
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchJobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){} 
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchJobType.innerText=arrSearchJobType[<%=session("language")%>]}catch(e){}
try{inner_SearchLineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{inner_SearchFactory.innerText=arrSearchFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchFamily.innerText=arrSearchFamily[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrSearchOptionFactory[<%=session("language")%>]}catch(e){}
try{inner_SearchStatus.innerText=arrSearchStatus[<%=session("language")%>]}catch(e){}
try{document.all.status.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{document.all.jobtype.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{document.all.jobstatus.options[0].text=arrSearchOptionStatus[<%=session("language")%>]}catch(e){}
try{inner_SearchCurrentStation.innerText=arrSearchCurrentStation[<%=session("language")%>]}catch(e){}
try{document.all.currentstation.options[0].text=arrSearchOptionCurrentStation[<%=session("language")%>]}catch(e){}
try{inner_SearchJobStartTime.innerText=arrSearchJobStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchJobCloseTime.innerText=arrSearchJobCloseTime[<%=session("language")%>]}catch(e){}
try{inner_SearchCycleTime.innerText=arrSearchCycleTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom1.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_SearchTo1.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
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
try{inner_CurrentStations.innerText=arrCurrentStations[<%=session("language")%>]}catch(e){}
try{inner_Operators.innerText=arrOperators[<%=session("language")%>]}catch(e){}
try{inner_LineLost.innerText=arrLineLost[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
try{inner_HoldReport.innerText=arrHoldReport[<%=session("language")%>]}catch(e){}
try{inner_HoldTime.innerText=arrHoldTime[<%=session("language")%>]}catch(e){}
try{inner_HoldType.innerText=arrHoldType[<%=session("language")%>]}catch(e){}
try{inner_HoldReason.innerText=arrHoldReason[<%=session("language")%>]}catch(e){}
try{inner_2DCodeReport.innerText=arr2DCodeReport[<%=session("language")%>]}catch(e){}
try{inner_2DCode.innerText=arr2DCode[<%=session("language")%>]}catch(e){}
try{inner_linkTime.innerText=arrLinkTime[<%=session("language")%>]}catch(e){}
try{inner_HoldJob.innerText=arrHoldJob[<%=session("language")%>]}catch(e){}
try{td_2DCode.innerText=arr2DCode[<%=session("language")%>]}catch(e){}
try{td_JobNumber.innerText=arrSearchJobNumber[<%=session("language")%>]}catch(e){}
try{td_linkUser.innerText=arrLinkUser[<%=session("language")%>]}catch(e){}
try{td_linkTime.innerText=arrLinkTime[<%=session("language")%>]}catch(e){}
try{td_SheetNumber.innerText=arrSheetNumber[<%=session("language")%>]}catch(e){}
try{td_Model.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{td_LineName.innerText=arrSearchLineName[<%=session("language")%>]}catch(e){}
try{td_HoldTime.innerText=arrHoldTime[<%=session("language")%>]}catch(e){}
try{td_ReleaseTime.innerText=arrReleaseTIme[<%=session("language")%>]}catch(e){}
try{td_PSSTime.innerText=arrPSSTIme[<%=session("language")%>]}catch(e){}
try{td_HoldPerson.innerText=arrHoldPerson[<%=session("language")%>]}catch(e){}
try{td_ReleasePerson.innerText=arrReleasePerson[<%=session("language")%>]}catch(e){}
try{td_HoldType.innerText=arrHoldType[<%=session("language")%>]}catch(e){}
try{td_HoldReason.innerText=arrHoldReason[<%=session("language")%>]}catch(e){}
try{td_OpCode.innerText=arrOpCode[<%=session("language")%>]}catch(e){}
try{td_HoldStation.innerText=arrHoldStation[<%=session("language")%>]}catch(e){}

}
</script>