<script language="javascript">
var strSearch="Search Line|��ѯ�߱�"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Labour|����Ͷ���" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strYearIndex="Year Index|���" 
var arrYearIndex=strYearIndex.split("|")
var strWeekIndex="Week Index|�ܴ�" 
var arrWeekIndex=strWeekIndex.split("|")
var strLineName="Line Name|�߱�" 
var arrLineName=strLineName.split("|")
var strRealWork="Real Work|���ڹ�ʱ" 
var arrRealWork=strRealWork.split("|")
var strTransferredWork="Transferred Work|ת�ƹ�ʱ" 
var arrTransferredWork=strTransferredWork.split("|")
var strIndirectWork="Indirect Work|��ֱ�ӹ�ʱ" 
var arrIndirectWork=strIndirectWork.split("|")
function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_SearchLine.innerText=arrSearchLine[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_NO.innerText=arrNO[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_YearIndex.innerText=arrYearIndex[<%=session("language")%>]}catch(e){}
try{inner_WeekIndex.innerText=arrWeekIndex[<%=session("language")%>]}catch(e){}
try{inner_LineName.innerText=arrLineName[<%=session("language")%>]}catch(e){}
try{inner_RealWork.innerText=arrRealWork[<%=session("language")%>]}catch(e){}
try{inner_TransferredWork.innerText=arrTransferredWork[<%=session("language")%>]}catch(e){}
try{inner_IndirectWork.innerText=arrIndirectWork[<%=session("language")%>]}catch(e){}
}
</script>
