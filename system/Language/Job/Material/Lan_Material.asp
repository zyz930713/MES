<script language="javascript">
var strSearch="Search Records|������¼" 
var arrSearch=strSearch.split("|")
var strSearchActionValue="Action Value|��ֵ" 
var arrSearchActionValue=strSearchActionValue.split("|")
var strSearchStartTime="Start Time|��ʼʱ��" 
var arrSearchStartTime=strSearchStartTime.split("|")
var strSearchFrom="From|��" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Job Material|�����������" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strConvertToJobYield="Convert to Job Yield|ת������������" 
var arrConvertToJobYield=strConvertToJobYield.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|�ͺ�" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|�߱�" 
var arrLine=strLine.split("|")
var strStation="Station|վ��" 
var arrStation=strStation.split("|")
var strActionName="Action Name|��������" 
var arrActionName=strActionName.split("|")
var strValue="Action Value|��ֵ" 
var arrValue=strValue.split("|")
var strRelativeActionName="Relative Action Name|��ز�������" 
var arrRelativeActionName=strRelativeActionName.split("|")
var strRelativeValue="Relative Value|�����ֵ" 
var arrRelativeValue=strRelativeValue.split("|")
var strRecords="No Records|û�м�¼" 
var arrRecords=strRecords.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){}
try{inner_SearchActionValue.innerText=arrSearchActionValue[<%=session("language")%>]}catch(e){}
try{inner_SearchStartTime.innerText=arrSearchStartTime[<%=session("language")%>]}catch(e){}
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){}
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{document.all.convert.value=arrConvertToJobYield[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_Progress.innerText=arrProgress[<%=session("language")%>]}catch(e){}
try{inner_Planer.innerText=arrPlaner[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_ActionName.innerText=arrActionName[<%=session("language")%>]}catch(e){}
try{inner_Value.innerText=arrValue[<%=session("language")%>]}catch(e){}
try{inner_RelativeActionName.innerText=arrRelativeActionName[<%=session("language")%>]}catch(e){}
try{inner_RelativeValue.innerText=arrRelativeValue[<%=session("language")%>]}catch(e){}
try{inner_NoRecords.innerText=arrRecords[<%=session("language")%>]}catch(e){}
}
</script>