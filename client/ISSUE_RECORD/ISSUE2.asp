<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%'response.Expires=0
'response.CacheControl="no-cache"%>
<!--#include virtual="/WOCF/BOCF_Open.asp" -->


<%LM_USER=trim(request("LM_USER"))
BOX_ID=trim(request("BOX_ID"))

ISSUE_ID=trim(request("ISSUE_ID"))


txtComments=trim(request("txtComments"))
OPEN_TIME=now()

set rs=server.CreateObject("adodb.recordset")
sql= "select * from  ISSUE_RECORD where ISSUE_ID='"&ISSUE_ID&"'"

rs.Open sql,conn,1,3
if rs.bof and rs.eof then


conn.Execute("INSERT INTO ISSUE_RECORD (LM_USER,BOX_ID,ISSUE_ID,txtComments,OPEN_TIME) values ('" &LM_USER&"','" &BOX_ID&"','" &ISSUE_ID&"','" &txtComments&"','" &OPEN_TIME&"')")
response.Redirect ("ISSUE.asp")
else
response.Redirect ("ISSUE.asp")

end if


%>