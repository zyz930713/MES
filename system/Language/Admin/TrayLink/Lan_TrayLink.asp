<script language="javascript">
var strSearch="Search TrayLink|查询Tray关系"
var arrSearch=strSearch.split("|")

var strPartNumber="Part Number|型号"
var arrPartNumber=strPartNumber.split("|")

var strStationName="Station Name|站点名称"
var arrStationName=strStationName.split("|")

var strUser="User|用户"
var arrUser=strUser.split("|")

var strNewTrayLink="Add a New TrayLink|新增Tray关系"
var arrNewTrayLink=strNewTrayLink.split("|")

var strNo="No|序号"
var arrNo=strNo.split("|")

var strStationChineseName="Station Chinese Name|站点中文名称"
var arrStationChineseName=strStationChineseName.split("|")

var strAction="Action|动作"
var arrAction=strAction.split("|")

var strStationSequence="Station Sequence|站点序号"
var arrStationSequence=strStationSequence.split("|")

var strTrayType="Tray Type|Tray类型"
var arrTrayType=strTrayType.split("|")

var strTraySize="Tray Size|Tray尺寸"
var arrTraySize=strTraySize.split("|")

var strTrayLinkList="Browse TrayLink list|浏览Tray关系列表"
var arrTrayLinkList=strTrayLinkList.split("|")

var strNoRecord="NoRecord|无数据"
var arrNoRecord=strNoRecord.split("|")

function language()
{
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){} 
try{inner_StationName.innerText=arrStationName[<%=session("language")%>]}catch(e){} 
try{inner_Search.innerText=arrSearch[<%=session("language")%>]}catch(e){} 
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){} 
try{inner_NewTrayLink.innerText=arrNewTrayLink[<%=session("language")%>]}catch(e){} 
try{inner_No.innerText=arrNo[<%=session("language")%>]}catch(e){} 
try{inner_StationChineseName.innerText=arrStationChineseName[<%=session("language")%>]}catch(e){} 
try{inner_PartNumber1.innerText=arrPartNumber[<%=session("language")%>]}catch(e){} 
try{inner_StationName1.innerText=arrStationName[<%=session("language")%>]}catch(e){} 
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){} 
try{inner_StationSequence.innerText=arrStationSequence[<%=session("language")%>]}catch(e){} 
try{inner_TrayType.innerText=arrTrayType[<%=session("language")%>]}catch(e){} 
try{inner_TraySize.innerText=arrTraySize[<%=session("language")%>]}catch(e){} 
try{inner_TrayLinkList.innerText=arrTrayLinkList[<%=session("language")%>]}catch(e){} 
try{inner_NoRecord.innerText=arrNoRecord[<%=session("language")%>]}catch(e){} 
}
language()
</script>
