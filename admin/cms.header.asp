<%
'option explicit
response.Codepage = 65001
response.Charset = "utf-8"
session.Timeout = 90

response.Buffer = true
response.ExpiresAbsolute = Now() - 1 
response.Expires = 0 
response.CacheControl = "no-cache" 
response.AddHeader "Pragma", "No-Cache"

if session("admin") = "" then
	response.write("<script>top.location.href='login.asp'</script>")
	response.end()
else
	session("admin") = session("admin")
end if
%>
<!--#include file="../content/site.config.asp" -->
<!--#include file="../include/ide.class.asp" -->
<!--#include file="../include/db.class.asp" -->
<!--#include file="../include/log.class.asp" -->
<%db.path = "../" & CONFIG_DATA%>
<!--#include file="../include/site.class.asp" -->
<%
function redirect(sUrl)
	if e.get("redirect") <> "" Then
		sUrl = e.get("redirect")
		if e.get("page") <> "" and inStr(sUrl,"page") < 1 then			
			sUrl = sUrl  & "&page=" & e.get("page")
		end if		
		response.redirect(sUrl)
	else		
		response.redirect(sUrl)
	end if
end function
%>