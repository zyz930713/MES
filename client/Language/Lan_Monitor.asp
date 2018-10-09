<script language="javascript">
var strSearchList="Search NCMR List|查询异常单列表"
var arrSearchList=strSearchList.split("|")
var strNumber ="NCMR NO.|异常单编号" 
var arrNumber=strNumber.split("|")
var strKeyWord="Key Word|关键字" 
var arrKeyWord=strKeyWord.split("|")
var strLine="Family|家族" 
var arrLine=strLine.split("|")
var strCreateTime="Create Time|申请时间" 
var arrCreateTime=strCreateTime.split("|")
var strStatus="Status|状态" 
var arrStatus=strStatus.split("|")
var strBrowse="Browse List|浏览异常单列表" 
var arrBrowse=strBrowse.split("|")
var strUser="User|用户" 
var arrUser=strUser.split("|")
var strNum="NO.|序列" 
var arrNum=strNum.split("|")
var strKeyWord1="Key Word|关键字" 
var arrKeyWord1=strKeyWord1.split("|")
var strLine1="Family|家族" 
var arrLine1=strLine1.split("|")
var strCreateTime1="Create Time|申请时间" 
var arrCreateTime1=strCreateTime1.split("|")
var strStatus1="Status|状态" 
var arrStatus1=strStatus1.split("|")
var strStatus2="Current Status|目前状态" 
var arrStatus2=strStatus2.split("|")
var strNumber1 ="NCMR NO.|异常单编号" 
var arrNumber1=strNumber1.split("|")
var strCurrent="Current Person|目前操作人" 
var arrCurrent=strCurrent.split("|")
var strCloseTime="Close Time|关闭时间" 
var arrCloseTime=strCloseTime.split("|")
var strRecords="No Records|没有记录:" 
var arrRecords=strRecords.split("|")
var strGreen="Green:Finished;|绿色：完成；" 
var arrGreen=strGreen.split("|")
var strRed="Red:Doing;|红色：进行中；" 
var arrRed=strRed.split("|")
var strBlue="Blue:Waiting;|蓝色：未进行；" 
var arrBlue=strBlue.split("|")
var strCauseKey="Cause Key Word|原因分析关键字"
var arrCauseKey=strCauseKey.split("|")
var strReject="Rejected Ticket|被拒绝的NCMR"
var arrReject=strReject.split("|")
var strFollow="Follow Up Person|措施验证人"
var arrFollow=strFollow.split("|")
var strCreator="Creator|申请人"
var arrCreator=strCreator.split("|")
var strCurrentStep="Current Status|目前状态"
var arrCurrentStep=strCurrentStep.split("|")

var strStation="Station|工站" 
var arrStation=strStation.split("|")
var strRepeat="Is Repeat|是否重复" 
var arrRepeat=strRepeat.split("|")
var strCloseDate="Close Time|关闭时间" 
var arrCloseDate=strCloseDate.split("|")
var strOwner="Owner|负责人"
var arrOwner=strOwner.split("|")

function language()
{
try{inner_SearchList.innerText=arrSearchList[<%=session("language")%>]}catch(e){} 
try{inner_Number.innerText=arrNumber[<%=session("language")%>]}catch(e){} 
try{inner_KeyWord.innerText=arrKeyWord[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_CreateTime.innerText=arrCreateTime[<%=session("language")%>]}catch(e){}
try{inner_Status.innerText=arrStatus[<%=session("language")%>]}catch(e){}
try{inner_Browse.innerText=arrBrowse[<%=session("language")%>]}catch(e){}
try{inner_User.innerText=arrUser[<%=session("language")%>]}catch(e){}
try{inner_Num.innerText=arrNum[<%=session("language")%>]}catch(e){}
try{inner_Number1.innerText=arrNumber1[<%=session("language")%>]}catch(e){} 
try{inner_KeyWord1.innerText=arrKeyWord1[<%=session("language")%>]}catch(e){}
try{inner_Line1.innerText=arrLine1[<%=session("language")%>]}catch(e){}
try{inner_CreateTime1.innerText=arrCreateTime1[<%=session("language")%>]}catch(e){}
try{inner_Status1.innerText=arrStatus1[<%=session("language")%>]}catch(e){}
try{inner_Status2.innerText=arrStatus2[<%=session("language")%>]}catch(e){}
try{inner_Current.innerText=arrCurrent[<%=session("language")%>]}catch(e){}
try{inner_CloseTime.innerText=arrCloseTime[<%=session("language")%>]}catch(e){}
try{inner_Records.innerText=arrRecords[<%=session("language")%>]}catch(e){}
try{inner_Green.innerText=arrGreen[<%=session("language")%>]}catch(e){}
try{inner_Red.innerText=arrRed[<%=session("language")%>]}catch(e){}
try{inner_Blue.innerText=arrBlue[<%=session("language")%>]}catch(e){}
try{inner_CauseKey.innerText=arrCauseKey[<%=session("language")%>]}catch(e){} 
try{inner_Reject.innerText=arrReject[<%=session("language")%>]}catch(e){} 
try{inner_Follow.innerText=arrFollow[<%=session("language")%>]}catch(e){}
try{inner_Creator.innerText=arrCreator[<%=session("language")%>]}catch(e){} 
try{inner_CurrentStep.innerText=arrCurrentStep[<%=session("language")%>]}catch(e){} 
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_CloseDate.innerText=arrCloseDate[<%=session("language")%>]}catch(e){}
try{inner_Repeat.innerText=arrRepeat[<%=session("language")%>]}catch(e){}
try{inner_Owner.innerText=arrOwner[<%=session("language")%>]}catch(e){}
}
</script>