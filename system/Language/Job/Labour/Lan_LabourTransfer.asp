<script language="javascript">
var strSearch="Search Line|查询线别"
var arrSearch=strSearch.split("|")
var strSearchLine="Line Name|线别" 
var arrSearchLine=strSearchLine.split("|")
var strBrowse="Browse Labour Trabsfer|浏览劳动力转移" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strbyPart="Browse by Part|按型号" 
var arrbyPart=strbyPart.split("|")
var strNO="NO|序列" 
var arrNO=strNO.split("|")
var strAction="Action|动作" 
var arrAction=strAction.split("|")
var strYearIndex="Year Index|年度" 
var arrYearIndex=strYearIndex.split("|")
var strWeekIndex="Week Index|周次" 
var arrWeekIndex=strWeekIndex.split("|")
var strTransferType="Transfer Type|转移类型" 
var arrTransferType=strTransferType.split("|")
var strTransferHour="Transfer Hour|转移时间" 
var arrTransferHour=strTransferHour.split("|")
var strInputCode="Input|输入人" 
var arrInputCode=strInputCode.split("|")
var strUpdateTime="Update Time|更新时间" 
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
