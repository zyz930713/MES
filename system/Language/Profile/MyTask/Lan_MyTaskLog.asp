<script language="javascript">
var strSearch="Search Task Log|搜索任务历史记录"
var arrSearch=strSearch.split("|")
var strSearchUser="User|用户"
var arrSearchUser=strSearchUser.split("|")
var strSearchResult="Result|任务结果"
var arrSearchResult=strSearchResult.split("|")
var strSearchResultSelect="--Select--|--选择--"
var arrSearchResultSelect=strSearchResultSelect.split("|")
var strSearchResultOK="Success|成功"
var arrSearchResultOK=strSearchResultOK.split("|")
var strSearchResultFail="Fail|失败"
var arrSearchResultFail=strSearchResultFail.split("|")
var strSearchRuntime="Run Time|运行时间"
var arrSearchRuntime=strSearchRuntime.split("|")
var strSearchFrom="From|从"
var arrSearchFrom=strSearchFrom.split("|")
var strSearchTo="To|至"
var arrSearchTo=strSearchTo.split("|")
var strSearchButton="Search|搜索"
var arrSearchButton=strSearchButton.split("|")
var strBrowse="My Task Hitory|我的自定义任务历史记录"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|用户"
var arrBrowseUser=strBrowseUser.split("|")
var strNO="No|编号" 
var arrNO=strNO.split("|")
var strDelete="Delete|编号" 
var arrDelete=strDelete.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strTaskName="Task Name|任务名称" 
var arrTaskName=strTaskName.split("|")
var strTaskType="Task Type|任务类型" 
var arrTaskType=strTaskType.split("|")
var strResult="Result|运行结果" 
var arrResult=strResult.split("|")
var strRunTime="Run Time|运行时间" 
var arrRunTime=strRunTime.split("|")
var strRecievers="Recievers|收件人" 
var arrRecievers=strRecievers.split("|")
var strRunErrors="Run Errors|运行错误" 
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