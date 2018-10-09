<script language="javascript">
var strSearch="Search Routine|查询制程"
var arrSearch=strSearch.split("|")
var strSearchPartNumber="Routine Name|制程编号" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchPartRule="Routine Rule|制程规则" 
var arrSearchPartRule=strSearchPartRule.split("|")
var strBrowse="Browse Routine List|查看制程列表" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strAdd="Add a New Routine|新增制程" 
var arrAdd=strAdd.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strPartNumber="Routine Name|制程名称" 
var arrPartNumber=strPartNumber.split("|")
var strPartRule="Routine Rule|制程规则" 
var arrPartRule=strPartRule.split("|")
var strPriority="Priority|优先级" 
var arrPriority=strPriority.split("|")
var strLines="Lines|线别" 
var arrLines=strLines.split("|")
var strFactory="Factory|工厂" 
var arrFactory=strFactory.split("|")
var strSection="Section|生产区段" 
var arrSection=strSection.split("|")
var strStations="Included Stations|站别" 
var arrStations=strStations.split("|")
var strTransactions="Station Transactions|站处理" 
var arrTransactions=strTransactions.split("|")
var strRoutine="Stations Routine|站的顺序" 
var arrRoutine=strRoutine.split("|")
var strYield="Target Yield|目标良率" 
var arrYield=strYield.split("|")
var strInterval="Max Interval|最大间隔时间" 
var arrInterval=strInterval.split("|")
var strPartType="Routing Type|类型" 
var arrPartType=strPartType.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchPartNumber.innerText=arrSearchPartNumber[<%=session("language")%>]}catch(e){}
try{inner_SearchPartRule.innerText=arrSearchPartRule[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_Add.innerText=arrAdd[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){}
try{inner_PartRule.innerText=arrPartRule[<%=session("language")%>]}catch(e){}
try{inner_Priority.innerText=arrPriority[<%=session("language")%>]}catch(e){}
try{inner_Lines.innerText=arrLines[<%=session("language")%>]}catch(e){}
try{document.all.factory.options[0].text=arrFactory[<%=session("language")%>]}catch(e){}
try{inner_Section.innerText=arrSection[<%=session("language")%>]}catch(e){}
try{inner_Stations.innerText=arrStations[<%=session("language")%>]}catch(e){}
try{inner_Transactions.innerText=arrTransactions[<%=session("language")%>]}catch(e){}
try{inner_Routine.innerText=arrRoutine[<%=session("language")%>]}catch(e){}
try{inner_Yield.innerText=arrYield[<%=session("language")%>]}catch(e){}
try{inner_Interval.innerText=arrInterval[<%=session("language")%>]}catch(e){}
try{inner_PartType.innerText=arrPartType[<%=session("language")%>]}catch(e){}
}
</script>
