<script language="javascript">
var strSearch="Search Line|��ѯ�߱�"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strSearchResult="Result|���" 
var arrSearchResult=strSearchResult.split("|")
var strBrowse="Browse Schedule Records|����ƻ��������м�¼" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|���ͺ�" 
var arrbyPart=strbyPart.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strStatus="Status|״̬" 
var arrStatus=strStatus.split("|")
var strLineName="Line Name|�߱�����" 
var arrLineName=strLineName.split("|")
var strPerson="Person|������" 
var arrPerson=strPerson.split("|")
var strShiftType="Shift Type|�������" 
var arrShiftType=strShiftType.split("|")
var strRuntime="Runtime|����ʱ��" 
var arrRuntime=strRuntime.split("|")
var strRunResult="Run Result|���н��" 
var arrRunResult=strRunResult.split("|")
var strErrorInfo="Error Info|������Ϣ" 
var arrErrorInfo=strErrorInfo.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_SearchResult.innerText=arrSearchResult[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_Person.innerText=arrPerson[<%=session("language")%>]}catch(e){}
try{inner_ShiftType.innerText=arrShiftType[<%=session("language")%>]}catch(e){}
try{inner_Runtime.innerText=arrRuntime[<%=session("language")%>]}catch(e){}
try{inner_RunResult.innerText=arrRunResult[<%=session("language")%>]}catch(e){}
try{inner_ErrorInfo.innerText=arrErrorInfo[<%=session("language")%>]}catch(e){}
}
</script>
