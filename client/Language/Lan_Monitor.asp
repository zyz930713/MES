<script language="javascript">
var strSearchList="Search NCMR List|��ѯ�쳣���б�"
var arrSearchList=strSearchList.split("|")
var strNumber ="NCMR NO.|�쳣�����" 
var arrNumber=strNumber.split("|")
var strKeyWord="Key Word|�ؼ���" 
var arrKeyWord=strKeyWord.split("|")
var strLine="Family|����" 
var arrLine=strLine.split("|")
var strCreateTime="Create Time|����ʱ��" 
var arrCreateTime=strCreateTime.split("|")
var strStatus="Status|״̬" 
var arrStatus=strStatus.split("|")
var strBrowse="Browse List|����쳣���б�" 
var arrBrowse=strBrowse.split("|")
var strUser="User|�û�" 
var arrUser=strUser.split("|")
var strNum="NO.|����" 
var arrNum=strNum.split("|")
var strKeyWord1="Key Word|�ؼ���" 
var arrKeyWord1=strKeyWord1.split("|")
var strLine1="Family|����" 
var arrLine1=strLine1.split("|")
var strCreateTime1="Create Time|����ʱ��" 
var arrCreateTime1=strCreateTime1.split("|")
var strStatus1="Status|״̬" 
var arrStatus1=strStatus1.split("|")
var strStatus2="Current Status|Ŀǰ״̬" 
var arrStatus2=strStatus2.split("|")
var strNumber1 ="NCMR NO.|�쳣�����" 
var arrNumber1=strNumber1.split("|")
var strCurrent="Current Person|Ŀǰ������" 
var arrCurrent=strCurrent.split("|")
var strCloseTime="Close Time|�ر�ʱ��" 
var arrCloseTime=strCloseTime.split("|")
var strRecords="No Records|û�м�¼:" 
var arrRecords=strRecords.split("|")
var strGreen="Green:Finished;|��ɫ����ɣ�" 
var arrGreen=strGreen.split("|")
var strRed="Red:Doing;|��ɫ�������У�" 
var arrRed=strRed.split("|")
var strBlue="Blue:Waiting;|��ɫ��δ���У�" 
var arrBlue=strBlue.split("|")
var strCauseKey="Cause Key Word|ԭ������ؼ���"
var arrCauseKey=strCauseKey.split("|")
var strReject="Rejected Ticket|���ܾ���NCMR"
var arrReject=strReject.split("|")
var strFollow="Follow Up Person|��ʩ��֤��"
var arrFollow=strFollow.split("|")
var strCreator="Creator|������"
var arrCreator=strCreator.split("|")
var strCurrentStep="Current Status|Ŀǰ״̬"
var arrCurrentStep=strCurrentStep.split("|")

var strStation="Station|��վ" 
var arrStation=strStation.split("|")
var strRepeat="Is Repeat|�Ƿ��ظ�" 
var arrRepeat=strRepeat.split("|")
var strCloseDate="Close Time|�ر�ʱ��" 
var arrCloseDate=strCloseDate.split("|")
var strOwner="Owner|������"
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