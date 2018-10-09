<%
function getrpt(rptname)
'base_info.rtpÊÇˆó±íŒ¦Ïó                             
Set rpt = Server.CreateObject("Crystal.CRPE.Application")                                                               
Path = Request.ServerVariables("PATH_TRANSLATED")                   
While (Right(Path, 1) <> "\" And Len(Path) <> 0)                      
iLen = Len(Path) - 1                                                  
Path = Left(Path, iLen)                                               
Wend 
rptfld=path&reportname
rpt.OpenReport(path& rptname, 1)
rpt.Options.MorePrintEngineErrorMessages = 0


selection_formula =cstr(field) & cstr(operator_words) & " " & "'"& cstr(value) &"'"
session("oRpt").DiscardSavedData
session("oRpt").RecordSelectionFormula =cstr( selection_formula)



On Error Resume Next                                                  
session("oRpt").ReadRecords                                           
If Err.Number <> 0 Then                                               
  Response.Write "An Error has occured on the server in attempting to access the data source"
Else

  If IsObject(session("oPageEngine")) Then                              
  	set session("oPageEngine") = nothing
  End If
set session("oPageEngine") = session("oRpt").PageEngine
end if                                         
end function
%>