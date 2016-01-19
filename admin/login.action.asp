<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
response.Codepage = 65001
response.Charset = "utf-8"
%>
<!--#include file="../content/site.config.asp" -->
<!--#include file="../include/ide.class.asp" -->
<!--#include file="../include/db.class.asp" -->
<%
db.path = "../" & CONFIG_DATA
%>
<!--#include file="../include/site.class.asp" -->
<!--#include file="../include/admin.class.asp" -->
<!--#include file="../include/log.class.asp" -->

<%
call site.safe()

if request.querystring = "login" then
	dim admin : set admin = new t_admin
		call admin.login()
else
	session.abandon()
	response.redirect("login.asp")
end if

set admin = nothing
set db = nothing
set e = nothing
%>

