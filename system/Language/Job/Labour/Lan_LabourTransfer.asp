<script language="javascript">
var strSearch="Search Line|��ѯ�߱�"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|�߱�" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Labour Trabsfer|����Ͷ���ת��" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|���ͺ�" 
var arrbyPart=strbyPart.split("|")
var strNO="NO|����" 
var arrNO=strNO.split("|")
var strAction="Action|����" 
var arrAction=strAction.split("|")
var strYearIndex="Year Index|���" 
var arrYearIndex=strYearIndex.split("|")
var strWeekIndex="Week Index|�ܴ�" 
var arrWeekIndex=strWeekIndex.split("|")
var strTransferType="Transfer Type|ת������" 
var arrTransferType=strTransferType.split("|")
var strTransferHour="Transfer Hour|ת��ʱ��" 
var arrTransferHour=strTransferHour.split("|")
var strInputCode="Input|������" 
var arrInputCode=strInputCode.split("|")
var strUpdateTime="Update Time|����ʱ��" 
var arrUpdateTime=strUpdateTime.split("|")
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
try{inner_TransferType.innerText=arrTransferType[<%=session("language")%>]}catch(e){}
try{inner_TransferHour.innerText=arrTransferHour[<%=session("language")%>]}catch(e){}
try{inner_InputCode.innerText=arrInputCode[<%=session("language")%>]}catch(e){}
try{inner_UpdateTime.innerText=arrUpdateTime[<%=session("language")%>]}catch(e){}
}
</script>
