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
var strBrowse="My Form Hitory|�ҵı���ʷ��¼"
var arrBrowse=strBrowse.split("|")
var strBrowseUser="User|�û�"
var arrBrowseUser=strBrowseUser.split("|")
var strNO="No|���" 
var arrNO=strNO.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strFormName="Form Name|������" 
var arrFormName=strFormName.split("|")
var strTransactCode="Transact Code|�����˹���" 
var arrTransactCode=strTransactCode.split("|")
var strTransactName="Transact Name|����������" 
var arrTransactName=strTransactName.split("|")
var strTransactResult="Transact Result|������" 
var arrTransactResult=strTransactResult.split("|")
var strTransactTime="Transact Time|����ʱ��" 
var arrTransactTime=strTransactTime.split("|")
var strRunErrors="Run Errors|���д���" 
var arrRunErrors=strRunErrors.split("|")
var strResultNote="Result Note|���ע��" 
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