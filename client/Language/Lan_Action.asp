<script language="javascript">
var strLine="Family|����" 
var arrLine=strLine.split("|")
var strTitle="NCMR Ticket(Short-term Action) |�쳣������¼�� ��������ʩ�� "
var arrTitle=strTitle.split("|")
var strFactory ="Factory|רҵ��" 
var arrFactory=strFactory.split("|")
var strKeyWord="Key Word|�ؼ���" 
var arrKeyWord=strKeyWord.split("|")
var strKeyWord1="Key Word|�ؼ���" 
var arrKeyWord1=strKeyWord1.split("|")
var strKeyWord2="Key Word|�ؼ���" 
var arrKeyWord2=strKeyWord2.split("|")
var strCreateTime="Create Time|����ʱ��" 
var arrCreateTime=strCreateTime.split("|")
var strCreator="Creator|������" 
var arrCreator=strCreator.split("|")
var strNum="NCMR NO.|�쳣�����" 
var arrNum=strNum.split("|")
var strJobNumber="Job Number|������" 
var arrJobNumber=strJobNumber.split("|")
var strModel="Model|�ͺ�" 
var arrModel=strModel.split("|")
var strStation="Station|��վ" 
var arrStation=strStation.split("|")
var strProblem="Problem Description|��������" 
var arrProblem=strProblem.split("|")
var strUpload="Upload File|�ϴ��ļ�" 
var arrUpload=strUpload.split("|")
var strOperator="Action Person|���δ�����" 
var arrOperator=strOperator.split("|")
var strAlert="(Please input your Badge Number)|(���������Ĺ���)" 
var arrAlert=strAlert.split("|")
var strNo ="NO.|���" 
var arrNo=strNo.split("|")
var strAction1="Short-term Action|������ʩ" 
var arrAction1=strAction1.split("|")
var strDepartment="Department|���β���" 
var arrDepartment=strDepartment.split("|")
var strAction="Action|��ȡ��ʩ" 
var arrAction=strAction.split("|")
var strPerson ="Person|������" 
var arrPerson=strPerson.split("|")
var strDueDate="Due Date|ָ���������" 
var arrDueDate=strDueDate.split("|")
var strDueDate1="Due Date|�������" 
var arrDueDate1=strDueDate1.split("|")
var strDueDate2="Due Date|�������" 
var arrDueDate2=strDueDate2.split("|")
var strDueDate3="Due Date|�������" 
var arrDueDate3=strDueDate3.split("|")
var strDueDate4="Due Date|�������" 
var arrDueDate4=strDueDate4.split("|")
var strRestart="Forward to:|ת��" 
var arrRestart=strRestart.split("|") 
var strDotice="Team Memebers|NCMR ��Ա" 
var arrDotice=strDotice.split("|")
var strSupervisor="Supervisor|����" 
var arrSupervisor=strSupervisor.split("|")
var strKeyPerson1="Defect Disposal Person|���ϸ��������" 
var arrKeyPerson1=strKeyPerson1.split("|")
var strKeyPerson2="Root Cause Person|ԭ����������" 
var arrKeyPerson2=strKeyPerson2.split("|")
var strKeyPerson3="Long-term Action Person|������ʩ�����" 
var arrKeyPerson3=strKeyPerson3.split("|")
var strKeyPerson4="QA|QA�����" 
var arrKeyPerson4=strKeyPerson4.split("|")
var strBackUp1="Backup|��ѡ�����" 
var arrBackUp1=strBackUp1.split("|")
var strBackUp2="Backup|��ѡ�����" 
var arrBackUp2=strBackUp2.split("|")
var strBackUp3="Backup|��ѡ�����" 
var arrBackUp3=strBackUp3.split("|")
var strBackUp4="Backup|��ѡ�����" 
var arrBackUp4=strBackUp4.split("|")
var strGreen="Green:Finished;|��ɫ����ɣ�" 
var arrGreen=strGreen.split("|")
var strRed="Red:Doing;|��ɫ�������У�" 
var arrRed=strRed.split("|")
var strBlue="Blue:Waiting;|��ɫ��δ���У�" 
var arrBlue=strBlue.split("|")
var strIsClosed="Is Closed|�˵��Ƿ�ر�" 
var arrIsClosed=strIsClosed.split("|")
var strCloseAlert="Please input reason if you want to close this NCMR|*�������Ϊ��NCMR�������ټ���������ȥ������дԭ��󣬵���ð�ť�رմ˵�"
var arrCloseAlert=strCloseAlert.split("|")
var strRisk="Risk Assessment|��������"
var arrRisk=strRisk.split("|")
var strDefect1="Defect Disposal|���ϸ���"
var arrDefect1=strDefect1.split("|")
var strDefect2="Reject|����"
var arrDefect2=strDefect2.split("|")
var strDefect3="Scrap|����"
var arrDefect3=strDefect3.split("|")
var strDefect4="Rework /Sorting|����/ѡ��"
var arrDefect4=strDefect4.split("|")
var strDefect5="Deviation|�������"
var arrDefect5=strDefect5.split("|")
var strDefect6="Others|����"
var arrDefect6=strDefect6.split("|")
var strRisk="Risk Assessment|��������"
var arrRisk=strRisk.split("|")
var strCause1="Root Cause|ԭ�����"
var arrCause1=strCause1.split("|")
var strCorrect1="Long-term Action|������ʩ"
var arrCorrect1=strCorrect1.split("|")
var strRejectHistory="Rejected History|�ܾ���ʷ"
var arrRejectHistory=strRejectHistory.split("|")
var strRejectStep="Rejected Step|�ܾ�����"
var arrRejectStep=strRejectStep.split("|")
var strRejectCause="Reason|�ܾ�ԭ��"
var arrRejectCause=strRejectCause.split("|")
var strRejectTime="Rejected Time|�ܾ�ʱ��"
var arrRejectTime=strRejectTime.split("|")
var strRejectPerson="Rejected Person|�ܾ���"
var arrRejectPerson=strRejectPerson.split("|")
var strRecords6="None|��"
var arrRecords6=strRecords6.split("|")

var strOwner="Owner|������"
var arrOwner=strOwner.split("|")
function language()
{
try{inner_Title.innerText=arrTitle[<%=session("language")%>]}catch(e){} 
try{inner_Factory.innerText=arrFactory[<%=session("language")%>]}catch(e){} 
try{inner_KeyWord.innerText=arrKeyWord[<%=session("language")%>]}catch(e){}
try{inner_KeyWord1.innerText=arrKeyWord1[<%=session("language")%>]}catch(e){}
try{inner_KeyWord2.innerText=arrKeyWord2[<%=session("language")%>]}catch(e){}
try{inner_Green.innerText=arrGreen[<%=session("language")%>]}catch(e){}
try{inner_Red.innerText=arrRed[<%=session("language")%>]}catch(e){}
try{inner_Blue.innerText=arrBlue[<%=session("language")%>]}catch(e){}
try{inner_Line.innerText=arrLine[<%=session("language")%>]}catch(e){}
try{inner_CreateTime.innerText=arrCreateTime[<%=session("language")%>]}catch(e){}
try{inner_Creator.innerText=arrCreator[<%=session("language")%>]}catch(e){}
try{inner_Num.innerText=arrNum[<%=session("language")%>]}catch(e){}
try{inner_JobNumber.innerText=arrJobNumber[<%=session("language")%>]}catch(e){}
try{inner_Model.innerText=arrModel[<%=session("language")%>]}catch(e){}
try{inner_Station.innerText=arrStation[<%=session("language")%>]}catch(e){}
try{inner_Problem.innerText=arrProblem[<%=session("language")%>]}catch(e){}
try{inner_Upload.innerText=arrUpload[<%=session("language")%>]}catch(e){}
try{inner_Operator.innerText=arrOperator[<%=session("language")%>]}catch(e){}
try{inner_Alert.innerText=arrAlert[<%=session("language")%>]}catch(e){}
try{inner_IsClosed.innerText=arrIsClosed[<%=session("language")%>]}catch(e){}
try{inner_CloseAlert.innerText=arrCloseAlert[<%=session("language")%>]}catch(e){}
try{inner_IsSkip.innerText=arrIsSkip[<%=session("language")%>]}catch(e){}
try{inner_IsSkip1.innerText=arrIsSkip1[<%=session("language")%>]}catch(e){}
try{inner_IsSkip2.innerText=arrIsSkip2[<%=session("language")%>]}catch(e){}
try{inner_IsSkip3.innerText=arrIsSkip3[<%=session("language")%>]}catch(e){}
try{inner_No.innerText=arrNo[<%=session("language")%>]}catch(e){}
try{inner_No1.innerText=arrNo[<%=session("language")%>]}catch(e){}
try{inner_Action1.innerText=arrAction1[<%=session("language")%>]}catch(e){}
try{inner_Action2.innerText=arrAction1[<%=session("language")%>]}catch(e){}
try{inner_DueDate1.innerText=arrDueDate1[<%=session("language")%>]}catch(e){}
try{inner_DueDate2.innerText=arrDueDate2[<%=session("language")%>]}catch(e){}
try{inner_DueDate3.innerText=arrDueDate3[<%=session("language")%>]}catch(e){}
try{inner_DueDate4.innerText=arrDueDate4[<%=session("language")%>]}catch(e){}
try{inner_DueDate.innerText=arrDueDate[<%=session("language")%>]}catch(e){}
try{inner_Department.innerText=arrDepartment[<%=session("language")%>]}catch(e){}
try{inner_Department2.innerText=arrDepartment[<%=session("language")%>]}catch(e){}
try{inner_Person.innerText=arrPerson[<%=session("language")%>]}catch(e){}
try{inner_Person2.innerText=arrPerson[<%=session("language")%>]}catch(e){}
try{inner_Action.innerText=arrAction[<%=session("language")%>]}catch(e){}
try{inner_Supervisor.innerText=arrSupervisor[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson1.innerText=arrKeyPerson1[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson2.innerText=arrKeyPerson2[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson3.innerText=arrKeyPerson3[<%=session("language")%>]}catch(e){}
try{inner_KeyPerson4.innerText=arrKeyPerson4[<%=session("language")%>]}catch(e){}
try{inner_BackUp1.innerText=arrBackUp1[<%=session("language")%>]}catch(e){}
try{inner_BackUp2.innerText=arrBackUp2[<%=session("language")%>]}catch(e){}
try{inner_BackUp3.innerText=arrBackUp3[<%=session("language")%>]}catch(e){}
try{inner_BackUp4.innerText=arrBackUp4[<%=session("language")%>]}catch(e){}
try{inner_Restart.innerText=arrRestart[<%=session("language")%>]}catch(e){}
try{inner_Dotice.innerText=arrDotice[<%=session("language")%>]}catch(e){}
try{inner_CloseAlert.innerText=arrCloseAlert[<%=session("language")%>]}catch(e){}
try{inner_Risk.innerText=arrRisk[<%=session("language")%>]}catch(e){}
try{inner_Defect1.innerText=arrDefect1[<%=session("language")%>]}catch(e){}
try{inner_Defect2.innerText=arrDefect2[<%=session("language")%>]}catch(e){}
try{inner_Defect3.innerText=arrDefect3[<%=session("language")%>]}catch(e){}
try{inner_Defect4.innerText=arrDefect4[<%=session("language")%>]}catch(e){}
try{inner_Defect5.innerText=arrDefect5[<%=session("language")%>]}catch(e){}
try{inner_Defect6.innerText=arrDefect6[<%=session("language")%>]}catch(e){}
try{inner_Risk.innerText=arrRisk[<%=session("language")%>]}catch(e){}
try{inner_Cause1.innerText=arrCause1[<%=session("language")%>]}catch(e){} 
try{inner_Correct1.innerText=arrCorrect1[<%=session("language")%>]}catch(e){} 
try{inner_RejectHistory.innerText=arrRejectHistory[<%=session("language")%>]}catch(e){} 
try{inner_RejectStep.innerText=arrRejectStep[<%=session("language")%>]}catch(e){} 
try{inner_RejectCause.innerText=arrRejectCause[<%=session("language")%>]}catch(e){} 
try{inner_RejectTime.innerText=arrRejectTime[<%=session("language")%>]}catch(e){} 
try{inner_RejectPerson.innerText=arrRejectPerson[<%=session("language")%>]}catch(e){} 
try{inner_Records6.innerText=arrRecords6[<%=session("language")%>]}catch(e){}
try{inner_Owner.innerText=arrOwner[<%=session("language")%>]}catch(e){}
}
</script>