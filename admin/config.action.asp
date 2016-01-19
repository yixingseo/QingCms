<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="cms.header.asp" -->
<!--#include file="../include/admin.class.asp" -->
<%
on error resume next
if request.form <> "" then
		
	for each config_name in request.form
		config_value = e.post(config_name)		
		if site.dic(config_name) <> e.post(config_name) then call save_item(config_name,config_value)
	next

	set comm = nothing
end if

sub save_item(config_name,config_value)
	if config_name = "" then exit sub
	if isnull(config_value) then config_value = ""
	set rs = e.objRs
	rs.open "select top 1 * from [t_config] where [config_name] = '"& config_name &"'",db.conn,1,3
	if rs.eof then
		rs.addnew()
		rs("config_name") = config_name
	end if	
	rs("config_value") = config_value
	rs.update
	rs.close
	set rs = nothing
end sub

if err then err.clear
redirect "site.config.asp"
%>
<!--#include file="cms.footer.asp" -->