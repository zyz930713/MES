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
var strBrowse="My Form Hitory|我的表单历史记录"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|用户"
var arrBrowseUser=strBrowseUser.split("|")
var strNO="No|编号" 
var arrNO=strNO.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strFormName="Form Name|表单名称" 
var arrFormName=strFormName.split("|")
var strTransactCode="Transact Code|处理人工号" 
var arrTransactCode=strTransactCode.split("|")
var strTransactName="Transact Name|处理人姓名" 
var arrTransactName=strTransactName.split("|")
var strTransactResult="Transact Result|处理结果" 
var arrTransactResult=strTransactResult.split("|")
var strTransactTime="Transact Time|处理时间" 
var arrTransactTime=strTransactTime.split("|")
var strRunErrors="Run Errors|运行错误" 
var arrRunErrors=strRunErrors.split("|")
var strResultNote="Result Note|结果注释" 
var arrResultNote=strResultNote.split("|")
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
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_FormName.innerText=arrFormName[<%=session("language")%>]}catch(e){}
try{inner_TransactCode.innerText=arrTransactCode[<%=session("language")%>]}catch(e){}
try{inner_TransactName.innerText=arrTransactName[<%=session("language")%>]}catch(e){}
try{inner_TransactResult.innerText=arrTransactResult[<%=session("language")%>]}catch(e){}
try{inner_TransactTime.innerText=arrTransactTime[<%=session("language")%>]}catch(e){}
try{inner_RunErrors.innerText=arrRunErrors[<%=session("language")%>]}catch(e){}
try{inner_ResultNote.innerText=arrResultNote[<%=session("language")%>]}catch(e){}
}
</script>