<script language="javascript">
var strSearch="Search TrayLink|��ѯTray��ϵ"
var arrSearch=strSearch.split("|")

var strPartNumber="Part Number|�ͺ�"
var arrPartNumber=strPartNumber.split("|")

var strStationName="Station Name|վ������"
var arrStationName=strStationName.split("|")

var strUser="User|�û�"
var arrUser=strUser.split("|")

var strNewTrayLink="Add a New TrayLink|����Tray��ϵ"
var arrNewTrayLink=strNewTrayLink.split("|")

var strNo="No|���"
var arrNo=strNo.split("|")

var strStationChineseName="Station Chinese Name|վ����������"
var arrStationChineseName=strStationChineseName.split("|")

var strAction="Action|����"
var arrAction=strAction.split("|")

var strStationSequence="Station Sequence|վ�����"
var arrStationSequence=strStationSequence.split("|")

var strTrayType="Tray Type|Tray����"
var arrTrayType=strTrayType.split("|")

var strTraySize="Tray Size|Tray�ߴ�"
var arrTraySize=strTraySize.split("|")

var strTrayLinkList="Browse TrayLink list|���Tray��ϵ�б�"
var arrTrayLinkList=strTrayLinkList.split("|")

var strNoRecord="NoRecord|������"
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
