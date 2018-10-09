<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%response.Expires=0
response.CacheControl="no-cache"%>
<!--#include virtual="/System/SystemCheck.asp" -->
<!--#include virtual="/WOCF/BOCF_Open.asp" -->
<%
path=request.Form("path")
query=request.Form("query")
query=replace(query,"*","&")
beforepath=path&"?"&query
trans=false
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
	trans=true
	exit for
	end if
next
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Trans.asp" -->
<%
end if
words=""
for i=1 to request.Form("idcount")
	if request.Form("id"&i)="1" then
		SQL="delete from SYSTEM_LOG where NID='"&request.Form("nid"&i)&"'"
		rs.open SQL,conn,1,3
	end if
next
if trans=true then
%>
<!--#include virtual="/WOCF/BOCF_Commit.asp" -->
<%end if%>
<!--#include virtual="/WOCF/BOCF_Close.asp" -->
<%
response.Redirect(beforepath)%>