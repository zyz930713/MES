<script language="javascript">
var strNewTrayLink="Add a New TrayLink|����Tray��ϵ"
var arrNewTrayLink=strNewTrayLink.split("|")

var strEditTrayLink="Edit a New TrayLink|�޸�Tray��ϵ"
var arrEditTrayLink=strEditTrayLink.split("|")

var strStationSequence="Station Sequence|վ�����"
var arrStationSequence=strStationSequence.split("|")

var strTrayType="Tray Type|Tray����"
var arrTrayType=strTrayType.split("|")

var strTraySize="Tray Size|Tray�ߴ�"
var arrTraySize=strTraySize.split("|")

var strPartNumber="Part Number|�ͺ�"
var arrPartNumber=strPartNumber.split("|")

var strStationName="Station Name|վ������"
var arrStationName=strStationName.split("|")
var arrBtnOK=[" OK ","ȷ��"]
var arrBtnReset=["Reset","����"]

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
