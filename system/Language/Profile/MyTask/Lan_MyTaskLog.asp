<script language="javascript">
var strSearch="Search Task Log|����������ʷ��¼"
var arrSearch=strSearch.split("|")
var strSearchUser="User|�û�"
var arrSearchUser=strSearchUser.split("|")
var strSearchResult="Result|������"
var arrSearchResult=strSearchResult.split("|")
var strSearchResultSelect="--Select--|--ѡ��--"
var arrSearchResultSelect=strSearchResultSelect.split("|")
var strSearchResultOK="Success|�ɹ�"
var arrSearchResultOK=strSearchResultOK.split("|")
var strSearchResultFail="Fail|ʧ��"
var arrSearchResultFail=strSearchResultFail.split("|")
var strSearchRuntime="Run Time|����ʱ��"
var arrSearchRuntime=strSearchRuntime.split("|")
var strSearchFrom="From|��"
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|��"
var arrSearchTo=strSearchTo.split("|")
var strSearchButton="Search|����"
var arrSearchButton=strSearchButton.split("|")
var strBrowse="My Task Hitory|�ҵ��Զ���������ʷ��¼"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|�û�"
var arrBrowseUser=strBrowseUser.split("|")
var strNO="No|���" 
var arrNO=strNO.split("|")
var strDelete="Delete|���" 
var arrDelete=strDelete.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strTaskName="Task Name|��������" 
var arrTaskName=strTaskName.split("|")
var strTaskType="Task Type|��������" 
var arrTaskType=strTaskType.split("|")
var strResult="Result|���н��" 
var arrResult=strResult.split("|")
var strRunTime="Run Time|����ʱ��" 
var arrRunTime=strRunTime.split("|")
var strRecievers="Recievers|�ռ���" 
var arrRecievers=strRecievers.split("|")
var strRunErrors="Run Errors|���д���" 
var arrRunErrors=strRunErrors.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchUser.innerText=arrSearchUser[<%=session("language")%>]}catch(e){} 
try{inner_SearchResult.innerText=arrSearchResult[<%=session("language")%>]}catch(e){} 
try{document.all.result.options[0].text=arrSearchResultSelect[<%=session("language")%>]}catch(e){} 
try{document.all.result.options[1].text=arrSearchResultOK[<%=session("language")%>]}catch(e){} 
try{document.all.result.options[2].text=arrSearchResultFail[<%=session("language")%>]}catch(e){} 
try{inner_SearchRuntime.innerText=arrSearchRuntime[<%=session("language")%>]}catch(e){} 
try{inner_SearchFrom.innerText=arrSearchFrom[<%=session("language")%>]}catch(e){} 
try{inner_SearchTo.innerText=arrSearchTo[<%=session("language")%>]}catch(e){} 
try{document.all.searchbutton.value=arrSearchButton[<%=session("language")%>]}catch(e){} 
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){} 
try{inner_BrowseUser.innerText=arrBrowseUser[<%=session("language")%>]}catch(e){} 
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){} 
try{inner_Delete.innerText=arrDelete[<%=session("language")%>]}catch(e){} 
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_TaskName.innerText=arrTaskName[<%=session("language")%>]}catch(e){}
try{inner_TaskType.innerText=arrTaskType[<%=session("language")%>]}catch(e){}
try{inner_Result.innerText=arrResult[<%=session("language")%>]}catch(e){}
try{inner_RunTime.innerText=arrRunTime[<%=session("language")%>]}catch(e){}
try{inner_Recievers.innerText=arrRecievers[<%=session("language")%>]}catch(e){}
try{inner_RunErrors.innerText=arrRunErrors[<%=session("language")%>]}catch(e){}
}
</script>