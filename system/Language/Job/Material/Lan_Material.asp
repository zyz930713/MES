<script language="javascript">
var strSearch="Search Records|搜索记录" 
var arrSearch=strSearch.split("|")
var strSearchActionValue="Action Value|数值" 
var arrSearchActionValue=strSearchActionValue.split("|")
var strSearchStartTime="Start Time|开始时间" 
var arrSearchStartTime=strSearchStartTime.split("|")
var strSearchFrom="From|从" 
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|到" 
var arrSearchTo=strSearchTo.split("|")
var strBrowse="Browse Job Material|浏览工单材料" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strConvertToJobYield="Convert to Job Yield|转换至工单良率" 
var arrConvertToJobYield=strConvertToJobYield.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strJobNumber="Job Number|工单号" 
var arrJobNumber=strJobNumber.split("|")
var strPartNumber="Part Number|型号" 
var arrPartNumber=strPartNumber.split("|")
var strLine="Line Name|线别" 
var arrLine=strLine.split("|")
var strStation="Station|站点" 
var arrStation=strStation.split("|")
var strActionName="Action Name|步骤名称" 
var arrActionName=strActionName.split("|")
var strValue="Action Value|数值" 
var arrValue=strValue.split("|")
var strRelativeActionName="Relative Action Name|相关步骤名称" 
var arrRelativeActionName=strRelativeActionName.split("|")
var strRelativeValue="Relative Value|相关数值" 
var arrRelativeValue=strRelativeValue.split("|")
var strRecords="No Records|没有记录" 
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