<script language="javascript">
var strSearch="Search Job|��ѯ����"
var arrSearch=strSearch.split("|")
var strSearchJobNumber="Job Number|������" 
var arrSearchJobNumber=strSearchJobNumber.split("|")
var strSearchPartNumber="Part Number|�ͺ�" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchStatus="Status|״̬" 
var arrSearchStatus=strSearchStatus.split("|")
var strSearchSheetNumber="Sheet Number|�ӹ������" 
var arrSearchSheetNumber=strSearchSheetNumber.split("|")
var strSearchLineName="LineName|�߱�" 
var arrSearchLineName=strSearchLineName.split("|")
var strSearchFactory="Factory|����" 
var arrSearchFactory=strSearchFactory.split("|")
var strSearchElapsedTime="Elapsed Time|��ʱ" 
var arrSearchElapsedTime=strSearchElapsedTime.split("|")
var strSearchM="m|����" 
var arrSearchM=strSearchM.split("|")
var strSearchJobStartTime="Job Start Time|������ʼʱ��" 
var arrSearchJobStartTime=strSearchJobStartTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="to|��" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Job list (Default in past 7 days)|��������б�Ĭ��7�����ڣ�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strStation="Station|վ��" 
var arrStation=strStation.split("|")
var strOperator="Operator|������" 
var arrOperator=strOperator.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strPartType="Part Type|�Ƴ�" 
var arrPartType=strPartType.split("|")
var strLine="Line|�߱�" 
var arrLine=strLine.split("|")
var strElapsedTime="Elapsed Time|��ʱ" 
var arrElapsedTime=strElapsedTime.split("|")
var strStartTime="Start Time|��ʼʱ��" 
var arrStartTime=strStartTime.split("|")
var strCloseTime="Close Time|����ʱ��" 
var arrCloseTime=strCloseTime.split("|")
var strRecords="No Records|û�м�¼" 
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