<script language="javascript">
var strSearch="Search Routine|��ѯ�Ƴ�"
var arrSearch=strSearch.split("|")
var strSearchPartNumber="Routine Name|�Ƴ̱��" 
var arrSearchPartNumber=strSearchPartNumber.split("|")
var strSearchPartRule="Routine Rule|�Ƴ̹���" 
var arrSearchPartRule=strSearchPartRule.split("|")
var strBrowse="Browse Routine List|�鿴�Ƴ��б�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strAdd="Add a New Routine|�����Ƴ�" 
var arrAdd=strAdd.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strPartNumber="Routine Name|�Ƴ�����" 
var arrPartNumber=strPartNumber.split("|")
var strPartRule="Routine Rule|�Ƴ̹���" 
var arrPartRule=strPartRule.split("|")
var strPriority="Priority|���ȼ�" 
var arrPriority=strPriority.split("|")
var strLines="Lines|�߱�" 
var arrLines=strLines.split("|")
var strFactory="Factory|����" 
var arrFactory=strFactory.split("|")
var strSection="Section|��������" 
var arrSection=strSection.split("|")
var strStations="Included Stations|վ��" 
var arrStations=strStations.split("|")
var strTransactions="Station Transactions|վ����" 
var arrTransactions=strTransactions.split("|")
var strRoutine="Stations Routine|վ��˳��" 
var arrRoutine=strRoutine.split("|")
var strYield="Target Yield|Ŀ������" 
var arrYield=strYield.split("|")
var strInterval="Max Interval|�����ʱ��" 
var arrInterval=strInterval.split("|")
var strPartType="Routing Type|����" 
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
