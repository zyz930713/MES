<script language="javascript">
var strNewTrayLink="Add a New TrayLink|新增Tray关系"
var arrNewTrayLink=strNewTrayLink.split("|")

var strEditTrayLink="Edit a New TrayLink|修改Tray关系"
var arrEditTrayLink=strEditTrayLink.split("|")

var strStationSequence="Station Sequence|站点序号"
var arrStationSequence=strStationSequence.split("|")

var strTrayType="Tray Type|Tray类型"
var arrTrayType=strTrayType.split("|")

var strTraySize="Tray Size|Tray尺寸"
var arrTraySize=strTraySize.split("|")

var strPartNumber="Part Number|型号"
var arrPartNumber=strPartNumber.split("|")

var strStationName="Station Name|站点名称"
var arrStationName=strStationName.split("|")
var arrBtnOK=[" OK ","确定"]
var arrBtnReset=["Reset","重置"]

function language()
{
try{inner_PartNumber.innerText=arrPartNumber[<%=session("language")%>]}catch(e){} 
try{inner_StationName.innerText=arrStationName[<%=session("language")%>]}catch(e){} 
try{inner_NewTrayLink.innerText=arrNewTrayLink[<%=session("language")%>]}catch(e){} 
try{inner_StationSequence.innerText=arrStationSequence[<%=session("language")%>]}catch(e){} 
try{inner_TrayType.innerText=arrTrayType[<%=session("language")%>]}catch(e){} 
try{inner_TraySize.innerText=arrTraySize[<%=session("language")%>]}catch(e){} 
try{inner_TrayLinkList.innerText=arrTrayLinkList[<%=session("language")%>]}catch(e){} 
try{inner_EditTrayLink.innerText=arrEditTrayLink[<%=session("language")%>]}catch(e){} 
try{document.all.btnOK.value=arrBtnOK[<%=session("language")%>]}catch(e){}
try{document.all.btnReset.value=arrBtnReset[<%=session("language")%>]}catch(e){}
}
language()
</script>
